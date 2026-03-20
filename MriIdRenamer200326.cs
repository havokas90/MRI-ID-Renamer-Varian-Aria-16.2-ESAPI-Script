// ================================================================
//  MriIdRenamer.cs
//  ESAPI Script for Varian Aria 16.2
//
//  PURPOSE:
//  When MRI studies are imported into Aria for fusion planning,
//  they arrive with placeholder image IDs (e.g. "1", "MRI_1").
//  This script automatically renames those IDs to something
//  meaningful, using information already stored in the database:
//    - The series description (series.Comment) tells us WHAT
//      the MRI sequence is (e.g. t2_tse_ax_p3)
//    - The study description (study.Comment) tells us WHERE
//      on the body the study was performed (e.g. Pelvis Prostate)
//
//  RESULT EXAMPLE:
//    Before : 1
//    After  : T2_TSE_AX_P3_PEL
//
//  SELECTIVE RENAMING:
//  The preview window shows all found MRI images with a checkbox
//  next to each one. Images are ticked by default. The user can
//  untick any image they do not want renamed — for example if
//  it was already named correctly in a previous course.
//  A "Select All / Deselect All" toggle is provided for convenience.
//
//  WHAT GETS RENAMED:
//  The script renames image.Id (the Volume Image ID), NOT
//  series.Id. This is because image.Id is the writeable property
//  confirmed to work in Aria, and it is what displays as the
//  image label in the Eclipse Image Registration workspace.
//
//  COMPATIBILITY:
//  Written without C# string interpolation ($"...") so it
//  compiles correctly on older Eclipse script engines that
//  use an earlier C# compiler version.
//
//  SAFETY:
//  - Only operates on the currently open patient (context.Patient)
//  - Only processes MR modality series
//  - Only renames 3D images (ZSize > 1), skipping 2D slices
//  - Shows a preview with checkboxes before making any changes
//  - Requires BeginModifications() and Eclipse Automation license
// ================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    // ============================================================
    //  SCRIPT ENTRY POINT
    //  Eclipse calls Execute() when the user runs the script.
    //  The ScriptContext object gives us access to the currently
    //  open patient, plan, image, and structure set.
    // ============================================================
    public class Script
    {
        public Script() { }

        [System.STAThread]
        public void Execute(ScriptContext context)
        {
            // ── Guard: patient must be open ──────────────────────────
            // context.Patient is null if no patient is open in Eclipse.
            // Always check this before doing anything else.
            if (context.Patient == null)
            {
                MessageBox.Show(
                    "Please open a patient in Aria before running this script.",
                    "No Patient Open",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // ── Collect all MRI images to rename ─────────────────────
            // We build a list of RenameItem objects, one per 3D MRI
            // image. This separates the data-gathering step from the
            // UI and the actual writing step.
            var items = new List<RenameItem>();

            // IMPORTANT: We call .ToList() on Studies and Series before
            // iterating. The ESAPI community confirms this is required
            // when you intend to modify objects during iteration.
            var studyList = context.Patient.Studies.ToList();

            foreach (var study in studyList)
            {
                // study.Comment is the Study Description — in Aria this
                // is usually set to something like "Pelvis Prostate" or
                // "Head Neck". We use this to determine the body region
                // suffix (e.g. _PEL, _HN). See GetBodySuffix() below.
                string studyDesc = (study.Comment ?? "").Trim();
                string suffix    = GetBodySuffix(studyDesc);

                var seriesList = study.Series.ToList();

                foreach (var series in seriesList)
                {
                    // Filter to MR modality only.
                    // series.Modality is a SeriesModality enum from
                    // VMS.TPS.Common.Model.Types. Using .ToString()
                    // gives us the DICOM string "MR", "CT", etc.
                    if (series.Modality.ToString() != "MR")
                        continue;

                    // series.Comment is the Series Description — the
                    // scanner's description of the sequence, e.g.
                    // "t2_tse_ax_p3_2.5mm". This is what we use as
                    // the basis for the new image ID.
                    string seriesDesc = (series.Comment ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(seriesDesc))
                        continue;

                    string proposedId = BuildNewId(seriesDesc, suffix);

                    // series.Images is IEnumerable<Image>.
                    // ToList() materialises it before we iterate.
                    var imageList = series.Images.ToList();

                    foreach (var image in imageList)
                    {
                        // image.ZSize is the number of slices.
                        // ZSize <= 1 means a 2D image (localiser, scout).
                        // We only rename the 3D volume used for fusion.
                        if (image.ZSize <= 1)
                            continue;

                        items.Add(new RenameItem
                        {
                            Image      = image,
                            CurrentId  = image.Id ?? "",
                            SeriesDesc = seriesDesc,
                            StudyDesc  = studyDesc,
                            Suffix     = suffix,
                            ProposedId = proposedId,
                            ZSize      = image.ZSize,

                            // Selected defaults to true — all images
                            // are ticked in the preview by default.
                            // The user can untick any they want to skip.
                            Selected   = true
                        });
                    }
                }
            }

            // ── Nothing found ────────────────────────────────────────
            if (items.Count == 0)
            {
                MessageBox.Show(
                    "No MRI series with descriptions and 3D images were found for this patient.\n\n" +
                    "Series must have a description (Comment) and contain a 3D image (ZSize > 1).",
                    "Nothing to Rename",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            // ── Launch the UI ────────────────────────────────────────
            var window = new RenamerWindow(context.Patient, items);
            window.ShowDialog();
        }

        // ============================================================
        //  GetBodySuffix
        //
        //  Reads the study description and returns a short body region
        //  suffix string, or empty string if no match is found.
        //
        //  HOW IT WORKS:
        //  The study description in Aria (study.Comment) typically
        //  contains plain English like "Pelvis Prostate" or "Head and
        //  Neck". We convert to uppercase and check for known keywords.
        //
        //  The checks are ordered deliberately:
        //    - "Head Neck" is checked before "Head" alone, so a study
        //      called "Head Neck" gets _HN rather than matching just
        //      the "Head" keyword below it.
        //    - "Abdomen" is checked before "Pelvis" so a combined
        //      "Abdomen Pelvis" study gets _AB rather than _PEL.
        //      Swap those two blocks if your department prefers pelvis
        //      to take priority.
        //
        //  TO ADD MORE REGIONS:
        //  Add a new if-block following the same pattern. The keyword
        //  array can hold as many strings as you need, including
        //  site-specific terms your department uses.
        // ============================================================
        public static string GetBodySuffix(string studyDescription)
        {
            if (string.IsNullOrWhiteSpace(studyDescription))
                return "";

            string upper = studyDescription.ToUpperInvariant();

            // Head & Neck — checked before individual Head/Neck entries
            if (ContainsAny(upper, new[] { "HEAD NECK", "HEAD AND NECK", "H&N", "HN" }))
                return "_HN";

            // Brain / Cranial
            if (ContainsAny(upper, new[] { "BRAIN", "CRANIAL", "SKULL" }))
                return "_BR";

            // Head (without explicit neck mention)
            if (ContainsAny(upper, new[] { "HEAD" }))
                return "_HN";

            // Neck (without explicit head mention)
            if (ContainsAny(upper, new[] { "NECK" }))
                return "_HN";

            // Chest / Thorax / Lung
            if (ContainsAny(upper, new[] { "CHEST", "THORAX", "THORACIC", "LUNG", "MEDIASTIN" }))
                return "_CH";

            // Abdomen — checked before Pelvis so "Abdomen Pelvis" gets _AB
            if (ContainsAny(upper, new[] { "ABDOMEN", "ABDOMINAL", "LIVER", "PANCREAS", "KIDNEY", "RENAL", "ADRENAL" }))
                return "_AB";

            // Pelvis and pelvic organs
            if (ContainsAny(upper, new[] { "PELVIS", "PELVIC", "PROSTATE", "BLADDER", "RECTUM", "UTERUS", "OVARY", "CERVIX", "VULVA" }))
                return "_PEL";

            // Spine
            if (ContainsAny(upper, new[] { "SPINE", "SPINAL", "LUMBAR", "SACRUM", "COCCYX", "CERVICAL SPINE", "THORACIC SPINE" }))
                return "_SP";

            // Extremities / musculoskeletal
            if (ContainsAny(upper, new[] { "LIMB", "LEG", "ARM", "KNEE", "HIP", "SHOULDER", "ELBOW", "WRIST", "ANKLE", "FEMUR", "TIBIA" }))
                return "_EXT";

            return "";
        }

        // ============================================================
        //  ContainsAny  (private helper)
        //  Returns true if the text contains ANY of the given keywords.
        //  Text should already be uppercased by the caller.
        // ============================================================
        private static bool ContainsAny(string text, string[] keywords)
        {
            foreach (var kw in keywords)
            {
                if (text.Contains(kw))
                    return true;
            }
            return false;
        }

        // ============================================================
        //  BuildNewId
        //
        //  Converts a series description and body suffix into a valid
        //  Aria image ID of at most 16 characters.
        //
        //  HOW IT WORKS:
        //
        //  Step 1 — Strip millimetre thickness patterns
        //  Series descriptions often end with slice thickness, e.g.
        //  "t2_tse_ax_p3_2.5mm". This is useful clinically but wastes
        //  characters in a 16-char ID. We remove patterns matching:
        //    underscore + digits + optional decimal + "mm"
        //  Examples removed: _2.5mm  _3mm  _0.5mm  _1mm  _1.5MM
        //
        //  Step 2 — Strip disallowed characters
        //  Aria IDs allow: letters, numbers, hyphens, underscores.
        //  Everything else is removed. Spaces are kept temporarily
        //  so we can use them as word boundaries in Step 4.
        //
        //  Step 3 — Calculate available space
        //  Total limit is 16 chars. We subtract the suffix length
        //  (e.g. 4 for "_PEL") to find space for the description.
        //
        //  Step 4 — Truncate at word boundary
        //  If the description is too long, we find the last space
        //  that fits within the limit and cut there. This avoids
        //  ugly mid-word cuts. Hard-cut if no word boundary found.
        //
        //  Step 5 — Finalise and append suffix
        //  Spaces become underscores, result is uppercased, and the
        //  body suffix is appended.
        //
        //  EXAMPLE:
        //    Input    = "t2_tse_ax_p3_2.5mm",  suffix = "_PEL"
        //    Step 1   = "t2_tse_ax_p3"         (removed _2.5mm)
        //    Step 3   = 16 - 4 = 12 chars available
        //    Step 4   = "t2_tse_ax_p3" fits in 12 exactly
        //    Step 5   = "T2_TSE_AX_P3" + "_PEL" = "T2_TSE_AX_P3_PEL"
        //    Length   = 16 ✓
        // ============================================================
        public static string BuildNewId(string seriesDescription, string suffix)
        {
            if (string.IsNullOrWhiteSpace(seriesDescription))
                return "MR_SERIES" + suffix;

            // Step 1: Remove millimetre thickness patterns.
            // Regex: _ + one or more digits + optional decimal + digits + mm
            string stripped = Regex.Replace(
                seriesDescription,
                @"_\d+\.?\d*mm",
                "",
                RegexOptions.IgnoreCase);

            // Step 2: Strip characters not allowed in Aria IDs.
            string clean = Regex.Replace(stripped.Trim(), @"[^A-Za-z0-9 \-_]", "").Trim();
            clean = Regex.Replace(clean, @"\s+", " ");

            if (string.IsNullOrWhiteSpace(clean))
                return "MR_SERIES" + suffix;

            // Step 3: Available characters for the description part.
            int available = 16 - suffix.Length;
            if (available < 1) available = 1;

            // Step 4: Truncate if needed.
            string descPart;
            if (clean.Length <= available)
            {
                descPart = clean;
            }
            else
            {
                int lastSpace = clean.LastIndexOf(' ', available - 1);
                if (lastSpace > 0)
                    descPart = clean.Substring(0, lastSpace).TrimEnd();
                else
                    descPart = clean.Substring(0, available);
            }

            // Step 5: Finalise and append suffix.
            string result = Finalise(descPart) + suffix;

            // Safety trim — should not normally be needed.
            if (result.Length > 16)
                result = result.Substring(0, 16);

            return result;
        }

        // ============================================================
        //  Finalise  (private helper)
        //  Replaces spaces with underscores, uppercases, trims trailing
        //  underscores that may have been left by a trailing space.
        // ============================================================
        private static string Finalise(string s)
        {
            return s.Replace(' ', '_').ToUpperInvariant().TrimEnd('_');
        }
    }

    // ================================================================
    //  RenameItem
    //
    //  Data container (DTO) for one rename operation.
    //  The Selected property is new — it tracks whether the user
    //  has ticked this item in the preview list. Only items with
    //  Selected == true will be renamed when Apply is clicked.
    // ================================================================
    public class RenameItem
    {
        public VMS.TPS.Common.Model.API.Image Image      { get; set; }
        public string CurrentId  { get; set; }
        public string SeriesDesc { get; set; }
        public string StudyDesc  { get; set; }
        public string Suffix     { get; set; }
        public string ProposedId { get; set; }
        public int    ZSize      { get; set; }

        // Whether this image is ticked for renaming.
        // Defaults to true so everything is selected initially.
        public bool Selected { get; set; }
    }

    // ================================================================
    //  RenamerWindow
    //
    //  WPF Window showing the preview list with checkboxes.
    //
    //  CHECKBOX BEHAVIOUR:
    //  Each row has a CheckBox on the left. Checking/unchecking it
    //  sets item.Selected and updates the Apply button text to show
    //  how many images are currently selected. This gives the user
    //  immediate feedback as they make selections.
    //
    //  The "Select All / Deselect All" toggle at the top of the list
    //  lets users quickly check or uncheck everything at once —
    //  useful if they want to start with nothing selected and only
    //  tick the new series.
    //
    //  The Apply button is disabled if nothing is selected, so the
    //  user cannot click Apply with zero items ticked.
    // ================================================================
    public class RenamerWindow : Window
    {
        private readonly Patient          _patient;
        private readonly List<RenameItem> _items;

        private ProgressBar _progress;
        private TextBlock   _status;
        private Button      _applyBtn;
        private CheckBox    _selectAllBox;

        // We keep a reference to each row's CheckBox so we can
        // update them programmatically when "Select All" is toggled.
        private readonly List<CheckBox> _rowCheckBoxes = new List<CheckBox>();

        // Prevents the SelectAll checkbox handler from firing while
        // we are programmatically updating individual row checkboxes.
        private bool _updatingSelectAll = false;

        public RenamerWindow(Patient patient, List<RenameItem> items)
        {
            _patient = patient;
            _items   = items;
            BuildUI();
        }

        private void BuildUI()
        {
            Title                 = "MRI ID Renamer — " + _patient.Id;
            Width                 = 760;
            SizeToContent         = SizeToContent.Height;
            ResizeMode            = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            var root = new StackPanel { Margin = new Thickness(18) };

            // Patient header
            root.Children.Add(MakeLabel(_patient.Name + "  (" + _patient.Id + ")", bold: true, size: 14));
            root.Children.Add(MakeLabel(_items.Count + " MRI image(s) found. Tick the ones you want to rename.", size: 11, gray: true));
            root.Children.Add(Gap(10));

            // ── Column headers ────────────────────────────────────────
            // Extra column on the left for the checkboxes.
            var headers = new Grid { Margin = new Thickness(0, 0, 0, 2) };
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(26) });  // Checkbox
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) }); // Current ID
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });  // Arrow
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(145) }); // New ID
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) });  // Suffix
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Desc

            // "Select All" checkbox in the header row.
            // Toggling it checks or unchecks all row checkboxes at once.
            _selectAllBox = new CheckBox
            {
                IsChecked           = true,  // all ticked by default
                VerticalAlignment   = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                ToolTip             = "Select / deselect all"
            };
            _selectAllBox.Checked   += OnSelectAllChanged;
            _selectAllBox.Unchecked += OnSelectAllChanged;
            AddToGrid(headers, _selectAllBox, 0);

            AddToGrid(headers, MakeLabel("Current ID",       bold: true, size: 11), 1);
            AddToGrid(headers, MakeLabel("",                 size: 11),             2);
            AddToGrid(headers, MakeLabel("New ID",           bold: true, size: 11), 3);
            AddToGrid(headers, MakeLabel("Suffix",           bold: true, size: 11), 4);
            AddToGrid(headers, MakeLabel("Series  |  Study", bold: true, size: 11), 5);
            root.Children.Add(headers);

            // ── Scrollable rename list ────────────────────────────────
            var listBorder = new Border
            {
                BorderBrush     = new SolidColorBrush(Colors.LightGray),
                BorderThickness = new Thickness(1),
                MaxHeight       = 260
            };
            var scroll    = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var listPanel = new StackPanel { Margin = new Thickness(6, 4, 6, 4) };

            foreach (var item in _items)
            {
                // Capture loop variable for use in the lambda below.
                // In C# foreach, the loop variable is shared across
                // iterations, so without this local copy the lambda
                // would always reference the last item in the list.
                var capturedItem = item;

                var row = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(26) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(145) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                // ── Per-row CheckBox ──────────────────────────────────
                // When the user ticks or unticks a row, we update
                // item.Selected and refresh the Apply button label.
                // We also update the SelectAll checkbox state to reflect
                // whether all, some, or none are selected.
                var cb = new CheckBox
                {
                    IsChecked           = item.Selected,
                    VerticalAlignment   = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                cb.Checked   += (s, e) => { capturedItem.Selected = true;  RefreshApplyButton(); UpdateSelectAllState(); };
                cb.Unchecked += (s, e) => { capturedItem.Selected = false; RefreshApplyButton(); UpdateSelectAllState(); };

                // Keep a reference so we can update it from SelectAll
                _rowCheckBoxes.Add(cb);
                AddToGrid(row, cb, 0);

                // Current ID
                AddToGrid(row, new TextBlock
                {
                    Text       = item.CurrentId,
                    Foreground = Brushes.Gray,
                    FontFamily = new FontFamily("Consolas"),
                    FontSize   = 11
                }, 1);

                // Arrow
                AddToGrid(row, new TextBlock
                {
                    Text                = "->",
                    HorizontalAlignment = HorizontalAlignment.Center,
                    FontSize            = 11
                }, 2);

                // Proposed new ID
                AddToGrid(row, new TextBlock
                {
                    Text       = item.ProposedId,
                    Foreground = new SolidColorBrush(Color.FromRgb(0, 128, 0)),
                    FontFamily = new FontFamily("Consolas"),
                    FontWeight = FontWeights.Bold,
                    FontSize   = 11
                }, 3);

                // Suffix — blue if matched, grey if not
                AddToGrid(row, new TextBlock
                {
                    Text       = string.IsNullOrEmpty(item.Suffix) ? "none" : item.Suffix,
                    Foreground = string.IsNullOrEmpty(item.Suffix)
                        ? Brushes.LightGray
                        : new SolidColorBrush(Color.FromRgb(0, 80, 160)),
                    FontFamily = new FontFamily("Consolas"),
                    FontWeight = FontWeights.Bold,
                    FontSize   = 11
                }, 4);

                // Series and study descriptions
                AddToGrid(row, new TextBlock
                {
                    Text         = item.SeriesDesc + "  |  " + item.StudyDesc,
                    Foreground   = Brushes.DimGray,
                    FontSize     = 10,
                    FontStyle    = FontStyles.Italic,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin       = new Thickness(4, 0, 0, 0)
                }, 5);

                listPanel.Children.Add(row);
            }

            scroll.Content   = listPanel;
            listBorder.Child = scroll;
            root.Children.Add(listBorder);
            root.Children.Add(Gap(14));

            // ── Progress bar ──────────────────────────────────────────
            _progress = new ProgressBar
            {
                Minimum    = 0,
                Maximum    = _items.Count,
                Value      = 0,
                Height     = 16,
                Visibility = Visibility.Hidden,
                Margin     = new Thickness(0, 0, 0, 8)
            };
            root.Children.Add(_progress);

            // ── Status / error message ────────────────────────────────
            _status = new TextBlock
            {
                Text         = "",
                FontSize     = 12,
                TextWrapping = TextWrapping.Wrap,
                MinHeight    = 18,
                Margin       = new Thickness(0, 0, 0, 10)
            };
            root.Children.Add(_status);

            // ── Apply button ──────────────────────────────────────────
            // Label shows the count of currently selected items.
            // Disabled if nothing is selected.
            _applyBtn = new Button
            {
                Content    = "Apply  (" + _items.Count + " selected)",
                Height     = 34,
                FontSize   = 13,
                FontWeight = FontWeights.SemiBold,
                Margin     = new Thickness(0, 0, 0, 6),
                IsEnabled  = _items.Count > 0
            };
            _applyBtn.Click += OnApply;
            root.Children.Add(_applyBtn);

            // ── Cancel button ─────────────────────────────────────────
            var cancelBtn = new Button { Content = "Cancel", Height = 28, FontSize = 12 };
            cancelBtn.Click += (s, e) => Close();
            root.Children.Add(cancelBtn);

            Content = root;
        }

        // ============================================================
        //  OnSelectAllChanged
        //  Called when the header CheckBox is ticked or unticked.
        //
        //  Sets _updatingSelectAll = true before touching the row
        //  checkboxes so that each row's Checked/Unchecked event does
        //  not call UpdateSelectAllState() in a loop — that would
        //  cause unnecessary recalculation on every single row.
        //  We update items and checkboxes together, then refresh once.
        // ============================================================
        private void OnSelectAllChanged(object sender, RoutedEventArgs e)
        {
            if (_updatingSelectAll) return;

            bool selectAll = _selectAllBox.IsChecked == true;

            _updatingSelectAll = true;
            for (int i = 0; i < _items.Count; i++)
            {
                _items[i].Selected    = selectAll;
                _rowCheckBoxes[i].IsChecked = selectAll;
            }
            _updatingSelectAll = false;

            RefreshApplyButton();
        }

        // ============================================================
        //  UpdateSelectAllState
        //  Called after an individual row checkbox changes.
        //
        //  Updates the header SelectAll checkbox to reflect reality:
        //    - All ticked   → SelectAll is checked
        //    - None ticked  → SelectAll is unchecked
        //    - Some ticked  → SelectAll is in indeterminate state (grey dash)
        //
        //  We set _updatingSelectAll = true first to prevent
        //  OnSelectAllChanged() from firing while we set IsChecked,
        //  which would incorrectly check/uncheck all rows.
        // ============================================================
        private void UpdateSelectAllState()
        {
            if (_updatingSelectAll) return;

            int selectedCount = _items.Count(i => i.Selected);

            _updatingSelectAll = true;
            if (selectedCount == _items.Count)
                _selectAllBox.IsChecked = true;
            else if (selectedCount == 0)
                _selectAllBox.IsChecked = false;
            else
                _selectAllBox.IsChecked = null; // null = indeterminate (grey dash)
            _updatingSelectAll = false;
        }

        // ============================================================
        //  RefreshApplyButton
        //  Updates the Apply button label with the current selected
        //  count and disables it if nothing is ticked.
        //  Called whenever a checkbox changes.
        // ============================================================
        private void RefreshApplyButton()
        {
            int count = _items.Count(i => i.Selected);
            _applyBtn.Content  = "Apply  (" + count + " selected)";
            _applyBtn.IsEnabled = count > 0;
        }

        // ============================================================
        //  OnApply
        //  Called when the user clicks Apply.
        //
        //  Only processes items where item.Selected == true.
        //  The progress bar Maximum is set to the selected count so
        //  the bar fills correctly regardless of how many were skipped.
        // ============================================================
        private void OnApply(object sender, RoutedEventArgs e)
        {
            var selectedItems = _items.Where(i => i.Selected).ToList();

            if (selectedItems.Count == 0) return;

            var confirm = MessageBox.Show(
                "Rename " + selectedItems.Count + " selected MRI image ID(s) for patient " + _patient.Id + "?\n\n" +
                "This cannot be undone from within this script.",
                "Confirm",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes) return;

            _applyBtn.IsEnabled  = false;
            _progress.Maximum    = selectedItems.Count;  // match progress bar to selected count
            _progress.Value      = 0;
            _progress.Visibility = Visibility.Visible;

            int ok = 0;
            var errors = new List<string>();

            try
            {
                // BeginModifications() must be called before any write.
                // It locks the patient record for editing. If this throws,
                // the system is not configured for write-enabled scripts.
                _patient.BeginModifications();

                foreach (var item in selectedItems)
                {
                    try
                    {
                        // THE ACTUAL RENAME:
                        // image.Id is the writeable property confirmed
                        // by the ESAPI community for renaming MRI images.
                        item.Image.Id = item.ProposedId;
                        ok++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add(item.CurrentId + ": " + ex.Message);
                    }

                    _progress.Value++;
                    DoEvents(); // keep UI responsive during loop
                }
            }
            catch (Exception ex)
            {
                SetStatus("Could not begin modifications: " + ex.Message, red: true);
                _applyBtn.IsEnabled  = true;
                _progress.Visibility = Visibility.Hidden;
                return;
            }

            // Report outcome
            if (errors.Count == 0)
            {
                SetStatus(ok + " image(s) renamed successfully.", red: false);
            }
            else
            {
                SetStatus(
                    ok + " renamed.   " + errors.Count + " failed:\n" +
                    string.Join("\n", errors.Select(err => "  - " + err)),
                    red: true);
            }

            // Repurpose Apply as Close now that the operation is done
            _applyBtn.Content   = "Close";
            _applyBtn.IsEnabled = true;
            _applyBtn.Click    -= OnApply;
            _applyBtn.Click    += (s, ev) => Close();
        }

        // ── Helper methods ────────────────────────────────────────────

        private void SetStatus(string msg, bool red)
        {
            _status.Text       = msg;
            _status.Foreground = red
                ? new SolidColorBrush(Colors.DarkRed)
                : new SolidColorBrush(Color.FromRgb(0, 128, 0));
        }

        // Pumps the WPF dispatcher so the progress bar repaints
        // during the synchronous rename loop.
        private static void DoEvents()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new Action(() => { }));
        }

        private static TextBlock MakeLabel(string text, bool bold = false,
            double size = 13, bool gray = false)
        {
            return new TextBlock
            {
                Text       = text,
                FontSize   = size,
                FontWeight = bold ? FontWeights.Bold : FontWeights.Normal,
                Foreground = gray ? Brushes.Gray : Brushes.Black
            };
        }

        private static void AddToGrid(Grid grid, UIElement element, int column)
        {
            Grid.SetColumn(element, column);
            grid.Children.Add(element);
        }

        private static UIElement Gap(double height)
        {
            return new Border { Height = height };
        }
    }
}
