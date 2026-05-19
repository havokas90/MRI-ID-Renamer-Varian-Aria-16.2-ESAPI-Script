using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyTitle("Eclipse Image Renamer")]
[assembly: AssemblyDescription("ARIA-hosted MR/CT image ID renaming script")]
[assembly: AssemblyProduct("Eclipse Image Renamer")]
[assembly: AssemblyCopyright("WCCD/CWCC 2026")]
[assembly: ComVisible(false)]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    /// <summary>
    /// ESAPI entry script.
    /// The Execute method is the hand-off point from Eclipse into our custom logic.
    /// </summary>
    public class Script
    {
        /// <summary>
        /// Build the preview list first, then hand it to a WPF dialog so the user can
        /// review and selectively apply the proposed image ID changes.
        /// </summary>
        public void Execute(ScriptContext context, Window window)
        {
            if (context == null)
            {
                MessageBox.Show("Script context was not provided.", "Eclipse Image Renamer", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Patient patient = context.Patient;
            if (patient == null)
            {
                MessageBox.Show("Open a patient in Eclipse before running Eclipse Image Renamer.", "Eclipse Image Renamer", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            List<RenameCandidate> candidates = BuildRenameCandidates(patient);
            if (candidates.Count == 0)
            {
                MessageBox.Show(
                    "No 3D MR or CT image volumes with a non-empty series description were found for the current patient.",
                    "Eclipse Image Renamer",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            EclipseImageRenamerWindow dialog = new EclipseImageRenamerWindow(patient, candidates);
            if (window != null)
            {
                dialog.Owner = window;
                dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
            {
                dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            dialog.ShowDialog();
        }

        private static List<RenameCandidate> BuildRenameCandidates(Patient patient)
        {
            List<RenameCandidate> candidates = new List<RenameCandidate>();
            Dictionary<string, int> usedIds = BuildCurrentIdCounts(patient);

            // The requested traversal order is patient -> studies -> series -> images,
            // and we deliberately materialise every ESAPI collection with ToList()
            // before iterating each level.
            foreach (Study study in patient.Studies.ToList())
            {
                string studyDescription = SafeTrim(study.Comment);
                string suffix = BuildBodySuffix(studyDescription);

                foreach (Series series in study.Series.ToList())
                {
                    string modality = SafeToUpper(series.Modality.ToString());
                    if (modality != "MR" && modality != "CT")
                    {
                        continue;
                    }

                    string seriesDescription = SafeTrim(series.Comment);
                    if (string.IsNullOrEmpty(seriesDescription))
                    {
                        continue;
                    }

                    foreach (VMS.TPS.Common.Model.API.Image image in series.Images.ToList())
                    {
                        // 2D scouts/localisers are filtered out by design. The user only
                        // wants renaming applied to 3D image volumes.
                        if (image.ZSize <= 1)
                        {
                            continue;
                        }

                        string currentId = SafeTrim(image.Id);

                        // Remove this image's current ID from the global "in use" pool
                        // before choosing a replacement. That allows the image to keep
                        // the same ID if that is already the best valid result.
                        AdjustCount(usedIds, currentId, -1);

                        string baseName = BuildBaseName(modality, seriesDescription);
                        string proposedId = BuildUniqueId(baseName, suffix, usedIds);

                        // Reserve the proposal immediately so later preview rows see it
                        // as already taken and get numbered consistently.
                        AdjustCount(usedIds, proposedId, 1);

                        RenameCandidate candidate = new RenameCandidate();
                        candidate.Image = image;
                        candidate.Modality = modality;
                        candidate.CurrentId = currentId;
                        candidate.NewId = proposedId;
                        candidate.Suffix = suffix;
                        candidate.SeriesDescription = seriesDescription;
                        candidate.StudyDescription = studyDescription;
                        candidate.IsSelected = true;
                        candidates.Add(candidate);
                    }
                }
            }

            return candidates;
        }

        private static Dictionary<string, int> BuildCurrentIdCounts(Patient patient)
        {
            // A dictionary of counts, rather than a simple set, lets us represent the
            // real patient state even if duplicate IDs already exist before the script
            // runs. That makes the uniqueness logic more robust.
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (Study study in patient.Studies.ToList())
            {
                foreach (Series series in study.Series.ToList())
                {
                    foreach (VMS.TPS.Common.Model.API.Image image in series.Images.ToList())
                    {
                        string id = SafeTrim(image.Id);
                        if (!string.IsNullOrEmpty(id))
                        {
                            AdjustCount(counts, id, 1);
                        }
                    }
                }
            }

            return counts;
        }

        private static void AdjustCount(Dictionary<string, int> counts, string key, int delta)
        {
            if (string.IsNullOrEmpty(key))
            {
                return;
            }

            int existing;
            counts.TryGetValue(key, out existing);
            existing = existing + delta;

            if (existing <= 0)
            {
                counts.Remove(key);
            }
            else
            {
                counts[key] = existing;
            }
        }

        private static string BuildBaseName(string modality, string rawDescription)
        {
            string description = SafeTrim(rawDescription);

            // Scanner descriptions often include slice-thickness tokens that are useful
            // in acquisition metadata but noisy inside a 16-character ARIA ID.
            description = Regex.Replace(description, "_\\d+\\.?\\d*mm", string.Empty, RegexOptions.IgnoreCase);
            description = CollapseSeparators(description);

            if (StartsWithIgnoreCase(description, "ep2d"))
            {
                string extracted = ExtractEp2dMeaningfulToken(description);
                if (!string.IsNullOrEmpty(extracted))
                {
                    description = extracted;
                }
            }

            description = Regex.Replace(description, "[^A-Za-z0-9_\\- ]", string.Empty);
            description = CollapseSeparators(description);

            if (string.IsNullOrEmpty(description))
            {
                description = modality;
            }

            if (modality == "CT")
            {
                // Only add our CT_ prefix when the scanner description has not already
                // started with CT_ or CT .
                if (!Regex.IsMatch(description, "^CT([_ ])", RegexOptions.IgnoreCase))
                {
                    description = "CT_" + description;
                }
            }

            return description;
        }

        private static string ExtractEp2dMeaningfulToken(string description)
        {
            string[] tokens = description.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return string.Empty;
            }

            string last = SafeTrim(tokens[tokens.Length - 1]);
            if (last.Length < 2)
            {
                return string.Empty;
            }

            if (Regex.IsMatch(last, "^\\d+$"))
            {
                return string.Empty;
            }

            return last;
        }

        private static string BuildUniqueId(string baseName, string suffix, Dictionary<string, int> usedIds)
        {
            // The first pass uses the plain name. If it collides, we append 2, 3, 4...
            // before the suffix and keep retrying until the ID is unique on the patient.
            int duplicateNumber = 1;

            while (true)
            {
                string candidate = ComposeFinalId(baseName, suffix, duplicateNumber);
                if (!usedIds.ContainsKey(candidate))
                {
                    return candidate;
                }

                duplicateNumber = duplicateNumber + 1;
            }
        }

        private static string ComposeFinalId(string baseName, string suffix, int duplicateNumber)
        {
            string safeSuffix = SafeTrim(suffix);
            string duplicateText = duplicateNumber > 1 ? duplicateNumber.ToString(CultureInfo.InvariantCulture) : string.Empty;
            int maxBaseLength = 16 - safeSuffix.Length - duplicateText.Length;

            if (maxBaseLength < 1)
            {
                maxBaseLength = 1;
            }

            string truncatedBase = TruncateBaseName(baseName, maxBaseLength);
            string combined = truncatedBase + duplicateText + safeSuffix;

            // Final output is always upper-case with spaces turned into underscores so
            // the value stored back into ARIA is consistent and easy to read.
            combined = combined.Replace(' ', '_').ToUpperInvariant();
            combined = Regex.Replace(combined, "_{2,}", "_");
            combined = combined.Trim('_', ' ');

            if (combined.Length > 16)
            {
                combined = combined.Substring(0, 16);
            }

            if (string.IsNullOrEmpty(combined))
            {
                combined = "IMAGE";
            }

            return combined;
        }

        private static string TruncateBaseName(string baseName, int maxLength)
        {
            string cleaned = SafeTrim(baseName);
            if (cleaned.Length <= maxLength)
            {
                return cleaned;
            }

            // Prefer cutting on an underscore boundary first, then a space boundary,
            // and only fall back to a hard substring when there is no cleaner option.
            string slice = cleaned.Substring(0, maxLength);
            int underscoreIndex = slice.LastIndexOf('_');
            if (underscoreIndex > 0)
            {
                return slice.Substring(0, underscoreIndex);
            }

            int spaceIndex = slice.LastIndexOf(' ');
            if (spaceIndex > 0)
            {
                return slice.Substring(0, spaceIndex);
            }

            return slice;
        }

        private static string BuildBodySuffix(string studyDescription)
        {
            string text = SafeTrim(studyDescription);
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            if (ContainsAny(text, new string[] { "Head Neck", "Head and Neck", "H&N", "HN", "Head", "Neck" }))
            {
                return "_HN";
            }

            if (ContainsAny(text, new string[] { "Brain", "Cranial", "Skull" }))
            {
                return "_BR";
            }

            if (ContainsAny(text, new string[] { "Chest", "Thorax", "Thoracic", "Lung", "Mediastin" }))
            {
                return "_CH";
            }

            if (ContainsAny(text, new string[] { "Abdomen", "Abdominal", "Liver", "Pancreas", "Kidney", "Renal", "Adrenal" }))
            {
                return "_AB";
            }

            if (ContainsAny(text, new string[] { "Pelvis", "Pelvic", "Prostate", "Bladder", "Rectum", "Uterus", "Ovary", "Cervix", "Vulva" }))
            {
                return "_PEL";
            }

            if (ContainsAny(text, new string[] { "Spine", "Spinal", "Lumbar", "Sacrum", "Coccyx", "Cervical Spine", "Thoracic Spine" }))
            {
                return "_SP";
            }

            if (ContainsAny(text, new string[] { "Limb", "Leg", "Arm", "Knee", "Hip", "Shoulder", "Elbow", "Wrist", "Ankle", "Femur", "Tibia" }))
            {
                return "_EXT";
            }

            return string.Empty;
        }

        private static bool ContainsAny(string text, string[] keywords)
        {
            foreach (string keyword in keywords)
            {
                if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool StartsWithIgnoreCase(string value, string prefix)
        {
            return value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static string CollapseSeparators(string value)
        {
            string result = SafeTrim(value);
            result = Regex.Replace(result, "\\s+", " ");
            result = Regex.Replace(result, "_{2,}", "_");
            result = result.Trim('_', ' ');
            return result;
        }

        private static string SafeTrim(string value)
        {
            return value == null ? string.Empty : value.Trim();
        }

        private static string SafeToUpper(string value)
        {
            return SafeTrim(value).ToUpperInvariant();
        }
    }

    public class RenameCandidate
    {
        public VMS.TPS.Common.Model.API.Image Image { get; set; }
        public string Modality { get; set; }
        public string CurrentId { get; set; }
        public string NewId { get; set; }
        public string Suffix { get; set; }
        public string SeriesDescription { get; set; }
        public string StudyDescription { get; set; }
        public bool IsSelected { get; set; }
        public CheckBox RowCheckBox { get; set; }
        public TextBox NewIdTextBox { get; set; }
        public Border RowBorder { get; set; }
    }

    /// <summary>
    /// Preview-and-apply dialog.
    /// The UI is hand-built in code so the script can stay in a single deployable file
    /// without any XAML dependency.
    /// </summary>
    public class EclipseImageRenamerWindow : Window
    {
        private readonly Patient _patient;
        private readonly ObservableCollection<RenameCandidate> _candidates;
        private readonly bool _previewOnly;
        private readonly Dictionary<string, int> _patientIdCounts;

        private CheckBox _selectAllCheckBox;
        private Button _applyButton;
        private Button _cancelButton;
        private ProgressBar _progressBar;
        private TextBlock _statusText;
        private TextBlock _countLabel;

        private bool _isUpdatingSelectAll;
        private bool _isNormalizingText;
        private bool _isApplying;
        private bool _isFinished;

        private readonly Brush _mrRowBrush = new SolidColorBrush(Color.FromRgb(232, 243, 252));
        private readonly Brush _ctRowBrush = new SolidColorBrush(Color.FromRgb(252, 238, 224));
        private readonly Brush _mrBadgeBrush = new SolidColorBrush(Color.FromRgb(88, 140, 190));
        private readonly Brush _ctBadgeBrush = new SolidColorBrush(Color.FromRgb(206, 128, 56));
        private readonly Brush _goodBrush = new SolidColorBrush(Color.FromRgb(46, 125, 50));
        private readonly Brush _suffixBrush = new SolidColorBrush(Color.FromRgb(37, 99, 235));
        private readonly Brush _mutedBrush = new SolidColorBrush(Color.FromRgb(130, 130, 130));
        private readonly Brush _errorBrush = new SolidColorBrush(Color.FromRgb(185, 28, 28));
        private readonly Brush _warningBrush = new SolidColorBrush(Color.FromRgb(180, 83, 9));
        private readonly Brush _defaultTextBoxBorderBrush = new SolidColorBrush(Color.FromRgb(171, 173, 179));

        public EclipseImageRenamerWindow(Patient patient, List<RenameCandidate> candidates)
            : this(patient, candidates, false)
        {
        }

        public EclipseImageRenamerWindow(List<RenameCandidate> candidates)
            : this(null, candidates, true)
        {
        }

        private EclipseImageRenamerWindow(Patient patient, List<RenameCandidate> candidates, bool previewOnly)
        {
            _patient = patient;
            _candidates = new ObservableCollection<RenameCandidate>(candidates);
            _previewOnly = previewOnly;
            _patientIdCounts = BuildPatientIdCounts(patient, candidates);

            Title = previewOnly ? "Eclipse Image Renamer Local Preview" : "Eclipse Image Renamer";
            Width = 1380;
            Height = 760;
            MinWidth = 1100;
            MinHeight = 600;
            ResizeMode = ResizeMode.CanResize;
            Background = Brushes.White;

            BuildUi();
            if (_previewOnly && _statusText != null)
            {
                _statusText.Text = "Local preview mode. No ESAPI patient is open and no image IDs will be written.";
            }
            UpdateSelectionSummary();
        }

        private void BuildUi()
        {
            // A DockPanel gives us a simple top / bottom / fill layout that works well
            // for a header, a scrolling preview table, and a persistent action footer.
            DockPanel root = new DockPanel();
            root.LastChildFill = true;
            Content = root;

            Border footer = BuildFooter();
            DockPanel.SetDock(footer, Dock.Bottom);
            root.Children.Add(footer);

            StackPanel topPanel = BuildHeader();
            DockPanel.SetDock(topPanel, Dock.Top);
            root.Children.Add(topPanel);

            ScrollViewer scrollViewer = new ScrollViewer();
            scrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            scrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
            scrollViewer.Margin = new Thickness(12, 0, 12, 12);
            root.Children.Add(scrollViewer);

            StackPanel listContainer = new StackPanel();
            scrollViewer.Content = listContainer;

            Border headerRow = new Border();
            headerRow.Background = new SolidColorBrush(Color.FromRgb(245, 245, 245));
            headerRow.BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220));
            headerRow.BorderThickness = new Thickness(1);
            headerRow.Padding = new Thickness(6);
            headerRow.Child = BuildGridRow(true, null);
            listContainer.Children.Add(headerRow);

            foreach (RenameCandidate candidate in _candidates)
            {
                Border rowBorder = new Border();
                rowBorder.Background = candidate.Modality == "MR" ? _mrRowBrush : _ctRowBrush;
                rowBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(235, 235, 235));
                rowBorder.BorderThickness = new Thickness(1, 0, 1, 1);
                rowBorder.Padding = new Thickness(6, 5, 6, 5);
                candidate.RowBorder = rowBorder;
                rowBorder.Child = BuildGridRow(false, candidate);
                listContainer.Children.Add(rowBorder);
            }
        }

        private StackPanel BuildHeader()
        {
            StackPanel container = new StackPanel();
            container.Margin = new Thickness(12, 12, 12, 10);

            TextBlock title = new TextBlock();
            title.Text = _previewOnly
                ? "Local preview: sample MR/CT image IDs"
                : "Patient: " + SafeText(_patient.Name) + "   |   ID: " + SafeText(_patient.Id);
            title.FontSize = 22;
            title.FontWeight = FontWeights.SemiBold;
            title.Foreground = new SolidColorBrush(Color.FromRgb(28, 28, 28));
            title.Margin = new Thickness(0, 0, 0, 8);
            container.Children.Add(title);

            int mrCount = _candidates.Count(c => c.Modality == "MR");
            int ctCount = _candidates.Count(c => c.Modality == "CT");

            _countLabel = new TextBlock();
            _countLabel.Text = mrCount.ToString(CultureInfo.InvariantCulture) + " MRI + " +
                               ctCount.ToString(CultureInfo.InvariantCulture) +
                               (_previewOnly
                                   ? " CT sample image(s). This preview cannot write to ARIA."
                                   : " CT image(s) found. Tick the ones you want to rename.");
            _countLabel.FontSize = 14;
            _countLabel.Margin = new Thickness(0, 0, 0, 8);
            container.Children.Add(_countLabel);

            StackPanel keyPanel = new StackPanel();
            keyPanel.Orientation = Orientation.Horizontal;
            keyPanel.Margin = new Thickness(0, 0, 0, 8);
            keyPanel.Children.Add(BuildColourKey("MRI", _mrRowBrush, _mrBadgeBrush));
            keyPanel.Children.Add(BuildColourKey("CT", _ctRowBrush, _ctBadgeBrush));
            container.Children.Add(keyPanel);

            return container;
        }

        private Border BuildColourKey(string label, Brush swatchBrush, Brush textBrush)
        {
            Border box = new Border();
            box.BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220));
            box.BorderThickness = new Thickness(1);
            box.CornerRadius = new CornerRadius(4);
            box.Background = Brushes.White;
            box.Padding = new Thickness(8, 4, 8, 4);
            box.Margin = new Thickness(0, 0, 10, 0);

            StackPanel inner = new StackPanel();
            inner.Orientation = Orientation.Horizontal;
            box.Child = inner;

            Border swatch = new Border();
            swatch.Width = 18;
            swatch.Height = 18;
            swatch.Background = swatchBrush;
            swatch.BorderBrush = textBrush;
            swatch.BorderThickness = new Thickness(1);
            swatch.Margin = new Thickness(0, 0, 6, 0);
            inner.Children.Add(swatch);

            TextBlock text = new TextBlock();
            text.Text = label;
            text.VerticalAlignment = VerticalAlignment.Center;
            text.Foreground = textBrush;
            text.FontWeight = FontWeights.Bold;
            inner.Children.Add(text);

            return box;
        }

        private Border BuildFooter()
        {
            Border footer = new Border();
            footer.BorderBrush = new SolidColorBrush(Color.FromRgb(225, 225, 225));
            footer.BorderThickness = new Thickness(0, 1, 0, 0);
            footer.Padding = new Thickness(12);
            footer.Background = new SolidColorBrush(Color.FromRgb(250, 250, 250));

            Grid footerGrid = new Grid();
            footerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            footerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition());
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            footer.Child = footerGrid;

            StackPanel statusPanel = new StackPanel();
            statusPanel.Orientation = Orientation.Vertical;
            Grid.SetRow(statusPanel, 0);
            Grid.SetColumn(statusPanel, 0);
            Grid.SetRowSpan(statusPanel, 2);
            footerGrid.Children.Add(statusPanel);

            _progressBar = new ProgressBar();
            _progressBar.Width = 380;
            _progressBar.Height = 16;
            _progressBar.Minimum = 0;
            _progressBar.Visibility = Visibility.Collapsed;
            _progressBar.Margin = new Thickness(0, 0, 0, 8);
            statusPanel.Children.Add(_progressBar);

            _statusText = new TextBlock();
            _statusText.Text = "Preview ready.";
            _statusText.Foreground = _mutedBrush;
            _statusText.TextWrapping = TextWrapping.Wrap;
            _statusText.MaxWidth = 850;
            statusPanel.Children.Add(_statusText);

            StackPanel buttonPanel = new StackPanel();
            buttonPanel.Orientation = Orientation.Horizontal;
            buttonPanel.HorizontalAlignment = HorizontalAlignment.Right;
            Grid.SetRow(buttonPanel, 0);
            Grid.SetColumn(buttonPanel, 1);
            footerGrid.Children.Add(buttonPanel);

            _applyButton = new Button();
            _applyButton.Content = "Apply (0 selected)";
            _applyButton.MinWidth = 155;
            _applyButton.Height = 34;
            _applyButton.Margin = new Thickness(0, 0, 8, 0);
            _applyButton.Click += ApplyButton_Click;
            buttonPanel.Children.Add(_applyButton);

            _cancelButton = new Button();
            _cancelButton.Content = "Cancel";
            _cancelButton.MinWidth = 95;
            _cancelButton.Height = 34;
            _cancelButton.Click += CancelButton_Click;
            buttonPanel.Children.Add(_cancelButton);

            return footer;
        }

        private Grid BuildGridRow(bool isHeader, RenameCandidate candidate)
        {
            // A plain Grid is used instead of DataGrid so the script has full control
            // over row colouring and cell contents while keeping the implementation
            // lightweight and single-file.
            Grid row = new Grid();
            row.VerticalAlignment = VerticalAlignment.Center;
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(44) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(72) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(42) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(260) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(260) });

            if (isHeader)
            {
                _selectAllCheckBox = new CheckBox();
                _selectAllCheckBox.IsThreeState = true;
                _selectAllCheckBox.VerticalAlignment = VerticalAlignment.Center;
                _selectAllCheckBox.HorizontalAlignment = HorizontalAlignment.Center;
                _selectAllCheckBox.Checked += SelectAllCheckBox_Changed;
                _selectAllCheckBox.Unchecked += SelectAllCheckBox_Changed;
                _selectAllCheckBox.Indeterminate += SelectAllCheckBox_Changed;
                AddCell(row, _selectAllCheckBox, 0);

                AddCell(row, BuildHeaderText("Mod"), 1);
                AddCell(row, BuildHeaderText("Current ID"), 2);
                AddCell(row, BuildHeaderText("->"), 3);
                AddCell(row, BuildHeaderText("New ID"), 4);
                AddCell(row, BuildHeaderText("Suffix"), 5);
                AddCell(row, BuildHeaderText("Series Description"), 6);
                AddCell(row, BuildHeaderText("Study Description"), 7);
                return row;
            }

            CheckBox box = new CheckBox();
            box.IsChecked = candidate.IsSelected;
            box.VerticalAlignment = VerticalAlignment.Center;
            box.HorizontalAlignment = HorizontalAlignment.Center;
            box.Tag = candidate;
            box.Checked += RowCheckBox_Changed;
            box.Unchecked += RowCheckBox_Changed;
            candidate.RowCheckBox = box;
            AddCell(row, box, 0);

            Border badge = new Border();
            badge.Background = candidate.Modality == "MR" ? _mrBadgeBrush : _ctBadgeBrush;
            badge.CornerRadius = new CornerRadius(10);
            badge.Padding = new Thickness(8, 3, 8, 3);
            badge.HorizontalAlignment = HorizontalAlignment.Left;
            badge.Child = new TextBlock
            {
                Text = candidate.Modality,
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                FontSize = 12
            };
            AddCell(row, badge, 1);

            AddCell(row, BuildBodyText(candidate.CurrentId), 2);
            AddCell(row, BuildBodyText("->"), 3);

            TextBox newIdBox = new TextBox();
            newIdBox.Text = SafeText(candidate.NewId);
            newIdBox.MaxLength = 16;
            newIdBox.CharacterCasing = CharacterCasing.Upper;
            newIdBox.FontWeight = FontWeights.Bold;
            newIdBox.Foreground = _goodBrush;
            newIdBox.VerticalAlignment = VerticalAlignment.Center;
            newIdBox.Tag = candidate;
            newIdBox.TextChanged += NewIdTextBox_TextChanged;
            candidate.NewIdTextBox = newIdBox;
            AddCell(row, newIdBox, 4);

            TextBlock suffixText = BuildBodyText(string.IsNullOrEmpty(candidate.Suffix) ? "none" : candidate.Suffix);
            suffixText.FontWeight = FontWeights.Bold;
            suffixText.Foreground = string.IsNullOrEmpty(candidate.Suffix) ? _mutedBrush : _suffixBrush;
            AddCell(row, suffixText, 5);

            AddCell(row, BuildWrappedText(candidate.SeriesDescription), 6);
            AddCell(row, BuildWrappedText(candidate.StudyDescription), 7);

            return row;
        }

        private TextBlock BuildHeaderText(string text)
        {
            TextBlock block = new TextBlock();
            block.Text = text;
            block.FontWeight = FontWeights.SemiBold;
            block.Foreground = new SolidColorBrush(Color.FromRgb(58, 58, 58));
            block.VerticalAlignment = VerticalAlignment.Center;
            return block;
        }

        private TextBlock BuildBodyText(string text)
        {
            TextBlock block = new TextBlock();
            block.Text = SafeText(text);
            block.Foreground = new SolidColorBrush(Color.FromRgb(28, 28, 28));
            block.VerticalAlignment = VerticalAlignment.Center;
            return block;
        }

        private TextBlock BuildWrappedText(string text)
        {
            TextBlock block = BuildBodyText(text);
            block.TextWrapping = TextWrapping.Wrap;
            return block;
        }

        private void AddCell(Grid row, UIElement element, int column)
        {
            Grid.SetColumn(element, column);
            row.Children.Add(element);
        }

        private void NewIdTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isNormalizingText)
            {
                return;
            }

            TextBox box = sender as TextBox;
            if (box == null)
            {
                return;
            }

            RenameCandidate candidate = box.Tag as RenameCandidate;
            if (candidate == null)
            {
                return;
            }

            string normalized = NormalizeEditedId(box.Text);
            if (box.Text != normalized)
            {
                int caretIndex = box.CaretIndex;
                _isNormalizingText = true;
                box.Text = normalized;
                box.CaretIndex = Math.Min(caretIndex, box.Text.Length);
                _isNormalizingText = false;
            }

            candidate.NewId = normalized;
            UpdateSelectionSummary();
        }

        private void RowCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            CheckBox box = sender as CheckBox;
            if (box == null)
            {
                return;
            }

            RenameCandidate candidate = box.Tag as RenameCandidate;
            if (candidate == null)
            {
                return;
            }

            candidate.IsSelected = box.IsChecked == true;
            UpdateSelectionSummary();
        }

        private void SelectAllCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingSelectAll || _selectAllCheckBox == null)
            {
                return;
            }

            bool? value = _selectAllCheckBox.IsChecked;
            if (!value.HasValue)
            {
                return;
            }

            foreach (RenameCandidate candidate in _candidates)
            {
                candidate.IsSelected = value.Value;
                if (candidate.RowCheckBox != null)
                {
                    candidate.RowCheckBox.IsChecked = value.Value;
                }
            }

            UpdateSelectionSummary();
        }

        private void UpdateSelectionSummary()
        {
            // This method is the single source of truth for selection-dependent UI:
            // it updates the Apply button label and drives the Select All tri-state.
            int selectedCount = _candidates.Count(c => c.IsSelected);
            int changedCount;
            string validationMessage;
            bool isValid = ValidateSelectionState(true, out validationMessage, out changedCount);

            if (_previewOnly)
            {
                _applyButton.Content = "Close Preview";
                _applyButton.IsEnabled = true;
            }
            else
            {
                _applyButton.Content = _isFinished
                    ? "Close"
                    : "Apply (" + changedCount.ToString(CultureInfo.InvariantCulture) + " changed)";
                _applyButton.IsEnabled = _isFinished || (changedCount > 0 && isValid);
            }

            if (!_isApplying && !_isFinished && _statusText != null)
            {
                if (!isValid)
                {
                    _statusText.Text = validationMessage;
                    _statusText.Foreground = _errorBrush;
                }
                else if (selectedCount > changedCount)
                {
                    _statusText.Text = changedCount.ToString(CultureInfo.InvariantCulture) +
                                       " selected image ID(s) will change. " +
                                       (selectedCount - changedCount).ToString(CultureInfo.InvariantCulture) +
                                       " unchanged selected row(s) will be skipped.";
                    _statusText.Foreground = _warningBrush;
                }
                else if (_previewOnly)
                {
                    _statusText.Text = "Local preview mode. No ESAPI patient is open and no image IDs will be written.";
                    _statusText.Foreground = _mutedBrush;
                }
                else
                {
                    _statusText.Text = selectedCount.ToString(CultureInfo.InvariantCulture) + " selected image ID(s) will change.";
                    _statusText.Foreground = _mutedBrush;
                }
            }

            _isUpdatingSelectAll = true;
            if (_selectAllCheckBox != null)
            {
                if (selectedCount == 0)
                {
                    _selectAllCheckBox.IsChecked = false;
                }
                else if (selectedCount == _candidates.Count)
                {
                    _selectAllCheckBox.IsChecked = true;
                }
                else
                {
                    _selectAllCheckBox.IsChecked = null;
                }
            }
            _isUpdatingSelectAll = false;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_previewOnly)
            {
                Close();
                return;
            }

            if (_isFinished)
            {
                Close();
                return;
            }

            List<RenameCandidate> selected = _candidates.Where(c => c.IsSelected).ToList();
            if (selected.Count == 0)
            {
                return;
            }

            int changedCount;
            string validationMessage;
            if (!ValidateSelectionState(true, out validationMessage, out changedCount))
            {
                _statusText.Text = validationMessage;
                _statusText.Foreground = _errorBrush;
                return;
            }

            List<RenameCandidate> toWrite = selected
                .Where(c => !IdsEqual(c.CurrentId, c.NewId))
                .ToList();

            if (toWrite.Count == 0)
            {
                _statusText.Text = "No image IDs changed. Unchanged selected rows were skipped.";
                _statusText.Foreground = _warningBrush;
                UpdateSelectionSummary();
                return;
            }

            if (!ConfirmApply(toWrite))
            {
                _statusText.Text = "Apply cancelled. No image IDs were changed.";
                _statusText.Foreground = _mutedBrush;
                UpdateSelectionSummary();
                return;
            }

            _isApplying = true;
            _progressBar.Visibility = Visibility.Visible;
            _progressBar.Minimum = 0;
            _progressBar.Maximum = toWrite.Count;
            _progressBar.Value = 0;
            _statusText.Text = "Applying image ID updates...";
            _statusText.Foreground = _mutedBrush;

            SetEditingEnabled(false);

            int successCount = 0;
            List<string> errors = new List<string>();

            try
            {
                // BeginModifications must happen before any writable ESAPI object is
                // changed. We do it once before the loop, then protect each item write
                // in its own try/catch so one bad image does not stop the batch.
                _patient.BeginModifications();

                for (int i = 0; i < toWrite.Count; i++)
                {
                    RenameCandidate candidate = toWrite[i];

                    try
                    {
                        candidate.Image.Id = candidate.NewId;
                        successCount = successCount + 1;
                    }
                    catch (Exception ex)
                    {
                        errors.Add(BuildErrorMessage(candidate, ex));
                    }

                    _progressBar.Value = i + 1;
                    DoEvents();
                }
            }
            catch (Exception ex)
            {
                errors.Add("Could not begin modifications for patient " + SafeText(_patient.Id) + ": " + ex.Message);
            }

            StringBuilder builder = new StringBuilder();
            builder.Append(successCount.ToString(CultureInfo.InvariantCulture));
            builder.Append(" image(s) renamed.");
            if (selected.Count > toWrite.Count)
            {
                builder.Append(" ");
                builder.Append((selected.Count - toWrite.Count).ToString(CultureInfo.InvariantCulture));
                builder.Append(" unchanged selected row(s) skipped.");
            }

            if (errors.Count > 0)
            {
                builder.Append(" ");
                builder.Append(errors.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append(" error(s): ");
                builder.Append(string.Join(" | ", errors.ToArray()));
                _statusText.Foreground = _errorBrush;
            }
            else
            {
                _statusText.Foreground = _goodBrush;
            }

            _statusText.Text = builder.ToString();
            _isApplying = false;
            _isFinished = true;
            UpdateSelectionSummary();
        }

        private static string BuildErrorMessage(RenameCandidate candidate, Exception ex)
        {
            string currentId = SafeText(candidate.CurrentId);
            string proposedId = SafeText(candidate.NewId);
            return currentId + " -> " + proposedId + " failed: " + ex.Message;
        }

        private void SetEditingEnabled(bool isEnabled)
        {
            _applyButton.IsEnabled = isEnabled;
            _selectAllCheckBox.IsEnabled = isEnabled;

            foreach (RenameCandidate candidate in _candidates)
            {
                if (candidate.RowCheckBox != null)
                {
                    candidate.RowCheckBox.IsEnabled = isEnabled;
                }

                if (candidate.NewIdTextBox != null)
                {
                    candidate.NewIdTextBox.IsEnabled = isEnabled;
                }
            }
        }

        private bool ConfirmApply(List<RenameCandidate> toWrite)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("You are about to rename ");
            builder.Append(toWrite.Count.ToString(CultureInfo.InvariantCulture));
            builder.Append(" image ID(s).");

            int previewCount = Math.Min(8, toWrite.Count);
            for (int i = 0; i < previewCount; i++)
            {
                builder.Append(Environment.NewLine);
                builder.Append(SafeText(toWrite[i].CurrentId));
                builder.Append(" -> ");
                builder.Append(SafeText(toWrite[i].NewId));
            }

            if (toWrite.Count > previewCount)
            {
                builder.Append(Environment.NewLine);
                builder.Append("...");
            }

            builder.Append(Environment.NewLine);
            builder.Append(Environment.NewLine);
            builder.Append("Continue?");

            return MessageBox.Show(
                       builder.ToString(),
                       "Confirm Eclipse Image Renamer changes",
                       MessageBoxButton.YesNo,
                       MessageBoxImage.Warning) == MessageBoxResult.Yes;
        }

        private bool ValidateSelectionState(bool updateUi, out string message, out int changedCount)
        {
            changedCount = 0;
            List<string> errors = new List<string>();
            Dictionary<string, int> proposedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<RenameCandidate>> proposedRows = new Dictionary<string, List<RenameCandidate>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<RenameCandidate, string> rowErrors = new Dictionary<RenameCandidate, string>();
            List<RenameCandidate> selected = _candidates.Where(c => c.IsSelected).ToList();
            Dictionary<string, int> remainingPatientIds = new Dictionary<string, int>(_patientIdCounts, StringComparer.OrdinalIgnoreCase);

            foreach (RenameCandidate candidate in selected)
            {
                string edited = candidate.NewIdTextBox != null ? candidate.NewIdTextBox.Text : candidate.NewId;
                string normalized = NormalizeEditedId(edited);
                candidate.NewId = normalized;

                if (candidate.NewIdTextBox != null && candidate.NewIdTextBox.Text != normalized)
                {
                    _isNormalizingText = true;
                    candidate.NewIdTextBox.Text = normalized;
                    candidate.NewIdTextBox.CaretIndex = candidate.NewIdTextBox.Text.Length;
                    _isNormalizingText = false;
                }

                if (string.IsNullOrEmpty(normalized))
                {
                    errors.Add(SafeText(candidate.CurrentId) + " has a blank New ID.");
                    rowErrors[candidate] = "New ID cannot be blank.";
                    continue;
                }

                if (IdsEqual(candidate.CurrentId, normalized))
                {
                    continue;
                }

                changedCount = changedCount + 1;
                AdjustValidationCount(remainingPatientIds, candidate.CurrentId, -1);

                int existing;
                proposedCounts.TryGetValue(normalized, out existing);
                proposedCounts[normalized] = existing + 1;

                List<RenameCandidate> rows;
                if (!proposedRows.TryGetValue(normalized, out rows))
                {
                    rows = new List<RenameCandidate>();
                    proposedRows[normalized] = rows;
                }
                rows.Add(candidate);
            }

            foreach (KeyValuePair<string, int> pair in proposedCounts)
            {
                if (pair.Value > 1)
                {
                    errors.Add("Duplicate New ID in selected rows: " + pair.Key + ".");
                    foreach (RenameCandidate row in proposedRows[pair.Key])
                    {
                        rowErrors[row] = "Duplicate selected New ID.";
                    }
                }
                else if (remainingPatientIds.ContainsKey(pair.Key))
                {
                    errors.Add("New ID already exists elsewhere in this patient: " + pair.Key + ".");
                    foreach (RenameCandidate row in proposedRows[pair.Key])
                    {
                        rowErrors[row] = "This ID already exists elsewhere in the patient.";
                    }
                }
            }

            if (updateUi)
            {
                foreach (RenameCandidate candidate in _candidates)
                {
                    string rowMessage;
                    rowErrors.TryGetValue(candidate, out rowMessage);
                    ApplyValidationStyle(candidate, rowMessage);
                }
            }

            if (errors.Count > 0)
            {
                message = "Fix the edited New ID values before applying: " + string.Join(" ", errors.ToArray());
                return false;
            }

            message = string.Empty;
            return true;
        }

        private void ApplyValidationStyle(RenameCandidate candidate, string rowMessage)
        {
            bool hasError = !string.IsNullOrEmpty(rowMessage);
            if (candidate.NewIdTextBox != null)
            {
                candidate.NewIdTextBox.BorderBrush = hasError ? _errorBrush : _defaultTextBoxBorderBrush;
                candidate.NewIdTextBox.BorderThickness = hasError ? new Thickness(2) : new Thickness(1);
                candidate.NewIdTextBox.ToolTip = hasError ? rowMessage : null;
            }

            if (candidate.RowBorder != null)
            {
                candidate.RowBorder.BorderBrush = hasError ? _errorBrush : new SolidColorBrush(Color.FromRgb(235, 235, 235));
                candidate.RowBorder.BorderThickness = hasError ? new Thickness(2) : new Thickness(1, 0, 1, 1);
            }
        }

        private static Dictionary<string, int> BuildPatientIdCounts(Patient patient, List<RenameCandidate> candidates)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (patient != null)
            {
                foreach (Study study in patient.Studies.ToList())
                {
                    foreach (Series series in study.Series.ToList())
                    {
                        foreach (VMS.TPS.Common.Model.API.Image image in series.Images.ToList())
                        {
                            AdjustValidationCount(counts, SafeText(image.Id), 1);
                        }
                    }
                }
            }
            else
            {
                foreach (RenameCandidate candidate in candidates)
                {
                    AdjustValidationCount(counts, candidate.CurrentId, 1);
                }
            }

            return counts;
        }

        private static void AdjustValidationCount(Dictionary<string, int> counts, string key, int delta)
        {
            string normalized = NormalizeEditedId(key);
            if (string.IsNullOrEmpty(normalized))
            {
                return;
            }

            int existing;
            counts.TryGetValue(normalized, out existing);
            existing = existing + delta;

            if (existing <= 0)
            {
                counts.Remove(normalized);
            }
            else
            {
                counts[normalized] = existing;
            }
        }

        private static bool IdsEqual(string left, string right)
        {
            return string.Equals(NormalizeEditedId(left), NormalizeEditedId(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeEditedId(string value)
        {
            string normalized = SafeText(value).Replace(' ', '_').ToUpperInvariant();
            normalized = Regex.Replace(normalized, "[^A-Z0-9_\\-]", string.Empty);
            normalized = Regex.Replace(normalized, "_{2,}", "_");
            normalized = normalized.Trim('_', ' ');

            if (normalized.Length > 16)
            {
                normalized = normalized.Substring(0, 16);
            }

            return normalized;
        }

        private static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new DispatcherOperationCallback(ExitFrame),
                frame);
            Dispatcher.PushFrame(frame);
        }

        private static object ExitFrame(object frameObject)
        {
            DispatcherFrame frame = frameObject as DispatcherFrame;
            if (frame != null)
            {
                frame.Continue = false;
            }

            return null;
        }

        private static string SafeText(string text)
        {
            return string.IsNullOrEmpty(text) ? string.Empty : text;
        }
    }
}
