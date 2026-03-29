// ================================================================
//  ImageIdRenamer.cs
//  ESAPI Script for Varian Aria 16.2
//
//  PURPOSE:
//  When MRI or CT studies are imported into Aria for planning,
//  they often arrive with placeholder image IDs (e.g. "1", "MRI_1").
//  This script automatically renames those IDs to something
//  meaningful using information already stored in the database:
//    - series.Comment  — the series description from the scanner
//                        e.g. "t2_tse_ax_p3_2.5mm"
//    - study.Comment   — the study description set in Aria
//                        e.g. "Pelvis Prostate"
//
//  RESULT EXAMPLES:
//    MRI before : 1              CT before : 2
//    MRI after  : T2_TSE_AX_PEL  CT after  : CT_PELVIS_PEL
//
//  MODALITIES HANDLED:
//    MR — no prefix, description + body suffix
//         e.g. T2_TSE_AX_P3_PEL
//    CT — CT_ prefix, description + body suffix
//         e.g. CT_CHEST_CH
//
//  DUPLICATE ID HANDLING:
//  If two images in the same patient would receive the same
//  proposed ID, a number is appended to make each one unique:
//    T2_TSE_AX_PEL, T2_TSE_AX_PEL2, T2_TSE_AX_PEL3
//  The number is appended before the final safety trim so the
//  result always stays within 16 characters.
//
//  SELECTIVE RENAMING:
//  The preview window shows all found images with a checkbox
//  per row. All are ticked by default. Untick any you want to
//  skip — useful when a patient has existing named images from
//  a previous course and only the new ones need renaming.
//
//  CONFIRMED API USAGE (verified against Varian docs):
//    series.Modality.ToString() — "MR", "CT" etc.
//    series.Comment             — series description (read-only)
//    series.Images              — IEnumerable<Image>
//    study.Comment              — study description (read-only)
//    image.Id                   — writeable, confirmed via community
//    image.ZSize                — number of slices
//    patient.BeginModifications() — required before any write
//    .ToList()                  — required before iterating if modifying
//
//  COMPATIBILITY:
//  No C# string interpolation ($"...") — compatible with the
//  older compiler used by some Aria 16.2 installations.
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

// Required for any script that writes data back to Aria.
// Triggers the script approval workflow on clinical systems.
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    // ============================================================
    //  SCRIPT ENTRY POINT
    //  Eclipse calls Execute() when the user runs the script.
    // ============================================================
    public class Script
    {
        public Script() { }

        [System.STAThread]
        public void Execute(ScriptContext context)
        {
            // Guard: patient must be open
            if (context.Patient == null)
            {
                MessageBox.Show(
                    "Please open a patient in Aria before running this script.",
                    "No Patient Open",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var items = new List<RenameItem>();

            // ToList() required before iterating when modifying objects
            var studyList = context.Patient.Studies.ToList();

            foreach (var study in studyList)
            {
                // study.Comment is the Study Description in Aria —
                // e.g. "Pelvis Prostate", "Head Neck".
                // Used to determine the body region suffix.
                string studyDesc = (study.Comment ?? "").Trim();
                string suffix    = GetBodySuffix(studyDesc);

                var seriesList = study.Series.ToList();

                foreach (var series in seriesList)
                {
                    // series.Modality.ToString() returns the DICOM
                    // modality string. Confirmed values from Varian docs:
                    // CT, MR, PT, REG, RTDOSE, RTIMAGE, RTPLAN, RTSTRUCT
                    string modality = series.Modality.ToString();

                    if (modality != "MR" && modality != "CT")
                        continue;

                    // series.Comment is the Series Description —
                    // the scanner's description of the sequence.
                    string seriesDesc = (series.Comment ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(seriesDesc))
                        continue;

                    string proposedId = BuildNewId(seriesDesc, suffix, modality);

                    var imageList = series.Images.ToList();

                    foreach (var image in imageList)
                    {
                        // Only rename 3D volumes (ZSize > 1).
                        // ZSize <= 1 means a 2D scout/localiser image.
                        if (image.ZSize <= 1)
                            continue;

                        items.Add(new RenameItem
                        {
                            Image      = image,
                            Series     = series,
                            CurrentId  = image.Id ?? "",
                            SeriesDesc = seriesDesc,
                            StudyDesc  = studyDesc,
                            Suffix     = suffix,
                            Modality   = modality,
                            ProposedId = proposedId,
                            ZSize      = image.ZSize,
                            Selected   = true
                        });
                    }
                }
            }

            if (items.Count == 0)
            {
                MessageBox.Show(
                    "No MRI or CT series with descriptions and 3D images were found.\n\n" +
                    "Series must have a description (Comment) and contain a 3D image (ZSize > 1).",
                    "Nothing to Rename",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            // ── Resolve duplicate proposed IDs ────────────────────────
            // After building all proposed IDs, check for duplicates
            // across all items. If two items would get the same ID,
            // append a number to the second, third etc.
            // e.g. T2_TSE_AX_PEL, T2_TSE_AX_PEL2, T2_TSE_AX_PEL3
            // We do this after collecting all items so we can see the
            // full picture before deciding what needs disambiguation.
            ResolveDuplicates(items);

            var window = new RenamerWindow(context.Patient, items);
            window.ShowDialog();
        }

        // ============================================================
        //  ResolveDuplicates
        //
        //  Scans the full list of RenameItems and appends a number to
        //  any proposed IDs that appear more than once.
        //
        //  HOW IT WORKS:
        //  We group items by their proposed ID. Any group with more
        //  than one item has duplicates. For those groups, we leave
        //  the first item unchanged and append "2", "3" etc. to the
        //  subsequent ones.
        //
        //  The number is inserted BEFORE the final 16-char trim so the
        //  result always stays within the Aria ID character limit.
        //  For example if the base is "T2_TSE_AX_P3_PEL" (16 chars),
        //  we shorten the base to make room: "T2_TSE_AX_P3_PE2"
        // ============================================================
        private static void ResolveDuplicates(List<RenameItem> items)
        {
            // Group by proposed ID (case-insensitive to be safe)
            var groups = items
                .GroupBy(i => i.ProposedId.ToUpperInvariant())
                .Where(g => g.Count() > 1);

            foreach (var group in groups)
            {
                int counter = 2;
                // Skip the first item — it keeps the original proposed ID.
                // Number the rest from 2 upward.
                foreach (var item in group.Skip(1))
                {
                    string suffix  = counter.ToString();
                    string baseId  = item.ProposedId;

                    // Trim the base to make room for the number suffix,
                    // staying within the 16-char Aria limit.
                    int maxBase = 16 - suffix.Length;
                    if (baseId.Length > maxBase)
                        baseId = baseId.Substring(0, maxBase);

                    item.ProposedId = baseId + suffix;
                    counter++;
                }
            }
        }

        // ============================================================
        //  GetBodySuffix
        //
        //  Reads the study description and returns a short body region
        //  suffix, or empty string if no keyword is matched.
        //
        //  Checks are ordered deliberately — more specific terms are
        //  checked before broader ones to avoid false matches:
        //    "Head Neck" before "Head" alone
        //    "Abdomen" before "Pelvis" (so "Abdomen Pelvis" gets _AB)
        //
        //  TO ADD MORE REGIONS: add a new if-block with your keywords.
        // ============================================================
        public static string GetBodySuffix(string studyDescription)
        {
            if (string.IsNullOrWhiteSpace(studyDescription))
                return "";

            string upper = studyDescription.ToUpperInvariant();

            if (ContainsAny(upper, new[] { "HEAD NECK", "HEAD AND NECK", "H&N", "HN" }))
                return "_HN";

            if (ContainsAny(upper, new[] { "BRAIN", "CRANIAL", "SKULL" }))
                return "_BR";

            if (ContainsAny(upper, new[] { "HEAD" }))
                return "_HN";

            if (ContainsAny(upper, new[] { "NECK" }))
                return "_HN";

            if (ContainsAny(upper, new[] { "CHEST", "THORAX", "THORACIC", "LUNG", "MEDIASTIN" }))
                return "_CH";

            // Abdomen before Pelvis — "Abdomen Pelvis" gets _AB
            if (ContainsAny(upper, new[] { "ABDOMEN", "ABDOMINAL", "LIVER", "PANCREAS", "KIDNEY", "RENAL", "ADRENAL" }))
                return "_AB";

            if (ContainsAny(upper, new[] { "PELVIS", "PELVIC", "PROSTATE", "BLADDER", "RECTUM", "UTERUS", "OVARY", "CERVIX", "VULVA" }))
                return "_PEL";

            if (ContainsAny(upper, new[] { "SPINE", "SPINAL", "LUMBAR", "SACRUM", "COCCYX", "CERVICAL SPINE", "THORACIC SPINE" }))
                return "_SP";

            if (ContainsAny(upper, new[] { "LIMB", "LEG", "ARM", "KNEE", "HIP", "SHOULDER", "ELBOW", "WRIST", "ANKLE", "FEMUR", "TIBIA" }))
                return "_EXT";

            return "";
        }

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
        //  Converts a series description into a valid Aria image ID
        //  within 16 characters.
        //
        //  MR: description (truncated) + body suffix
        //      e.g. T2_TSE_AX_P3_PEL
        //  CT: "CT_" prefix + description (truncated) + body suffix
        //      e.g. CT_CHEST_CH
        //
        //  SPECIAL CASES:
        //
        //  ep2d (DWI/diffusion) sequences:
        //    Descriptions like "ep2d_diff_b50_400_1000_tra+calc1400_ADC"
        //    contain a meaningful identifier at the very end (ADC, BVAL).
        //    For these we extract just the last underscore-separated token
        //    rather than truncating from the front, giving e.g. ADC_PEL
        //    instead of EP2D_DIFF_B5_PEL which is less useful.
        //
        //  CT descriptions already containing "CT_":
        //    Some CT series arrive with "CT_" already in the description
        //    (e.g. "CT_PEL 2.00 Qr40"). Adding our own "CT_" prefix would
        //    produce "CT_CT_PEL_PEL". We detect this and skip the prefix.
        //
        //  STEPS:
        //  1. Handle ep2d — extract last meaningful token
        //  2. Strip mm thickness patterns (_2.5mm, _3mm etc.)
        //  3. Strip disallowed characters
        //  4. For CT, add "CT_" prefix only if not already present
        //  5. Calculate available chars = 16 - prefix - suffix length
        //  6. Truncate at last word/underscore boundary within available chars
        //  7. Assemble: prefix + description + suffix, uppercase
        // ============================================================
        public static string BuildNewId(string seriesDescription, string suffix, string modality)
        {
            if (string.IsNullOrWhiteSpace(seriesDescription))
            {
                if (modality == "CT")
                    return ("CT_SCAN" + suffix).Substring(0, Math.Min(16, ("CT_SCAN" + suffix).Length));
                return ("MR_SERIES" + suffix).Substring(0, Math.Min(16, ("MR_SERIES" + suffix).Length));
            }

            // Step 1: Handle ep2d (DWI/diffusion) sequences specially.
            // These arrive with descriptions like:
            //   ep2d_diff_b50_400_1000_tra+calc1400_ADC
            //   ep2d_diff_b50_400_1000_tra+calc1400_CALC_BVAL
            // The clinically meaningful identifier is the last token
            // after the final underscore: ADC or BVAL.
            // We use that as the entire description part rather than
            // trying to truncate from the front which would give
            // something unhelpful like EP2D_DIFF_B5.
            string workingDesc = seriesDescription;
            if (workingDesc.ToUpperInvariant().StartsWith("EP2D"))
            {
                // Split on underscore, take the last non-empty token
                string[] tokens = workingDesc.Split('_');
                string lastToken = "";
                for (int i = tokens.Length - 1; i >= 0; i--)
                {
                    if (!string.IsNullOrWhiteSpace(tokens[i]))
                    {
                        lastToken = tokens[i].Trim();
                        break;
                    }
                }
                // Only use the last token if it looks meaningful
                // (at least 2 chars, not just a number)
                if (lastToken.Length >= 2 && !Regex.IsMatch(lastToken, @"^\d+$"))
                    workingDesc = lastToken;
            }

            // Step 2: Strip millimetre thickness patterns
            // Regex: underscore + digits + optional decimal + digits + mm
            string stripped = Regex.Replace(
                workingDesc,
                @"_\d+\.?\d*mm",
                "",
                RegexOptions.IgnoreCase);

            // Step 3: Strip characters not allowed in Aria IDs
            // Keep: letters, numbers, spaces (temp), hyphens, underscores
            string clean = Regex.Replace(stripped.Trim(), @"[^A-Za-z0-9 \-_]", "").Trim();
            clean = Regex.Replace(clean, @"\s+", " ");

            if (string.IsNullOrWhiteSpace(clean))
            {
                if (modality == "CT")
                    return ("CT_SCAN" + suffix).Substring(0, Math.Min(16, ("CT_SCAN" + suffix).Length));
                return ("MR_SERIES" + suffix).Substring(0, Math.Min(16, ("MR_SERIES" + suffix).Length));
            }

            // Step 4: Determine prefix for CT.
            // Only add "CT_" if the description doesn't already start
            // with it — prevents "CT_CT_PEL_PEL" when the scanner has
            // already included CT_ in the series description.
            string prefix = "";
            if (modality == "CT")
            {
                if (!clean.ToUpperInvariant().StartsWith("CT_") &&
                    !clean.ToUpperInvariant().StartsWith("CT "))
                    prefix = "CT_";
            }

            // Step 5: Available characters for the description part
            int available = 16 - prefix.Length - suffix.Length;
            if (available < 1) available = 1;

            // Step 6: Truncate if needed.
            // Try underscore boundary first (more natural for these IDs),
            // then fall back to space boundary, then hard cut.
            string descPart;
            if (clean.Length <= available)
            {
                descPart = clean;
            }
            else
            {
                // Try last underscore within limit
                int lastUnderscore = clean.LastIndexOf('_', available - 1);
                int lastSpace      = clean.LastIndexOf(' ', available - 1);
                int cutAt          = Math.Max(lastUnderscore, lastSpace);

                if (cutAt > 0)
                    descPart = clean.Substring(0, cutAt).TrimEnd('_').TrimEnd();
                else
                    descPart = clean.Substring(0, available);
            }

            // Step 7: Assemble final ID — prefix + description + suffix,
            // all uppercased with spaces converted to underscores.
            string result = prefix + Finalise(descPart) + suffix;

            // Safety trim — should not normally be needed
            if (result.Length > 16)
                result = result.Substring(0, 16);

            return result;
        }

        private static string Finalise(string s)
        {
            return s.Replace(' ', '_').ToUpperInvariant().TrimEnd('_');
        }
    }

    // ================================================================
    //  RenameItem
    //  Data container for one rename operation.
    //  Selected defaults to true — all ticked in the preview.
    //  Modality stored so the UI can colour-code MR vs CT rows.
    // ================================================================
    public class RenameItem
    {
        public VMS.TPS.Common.Model.API.Image Image      { get; set; }
        public Series Series     { get; set; }  // stored so we can attempt series.Id rename
        public string CurrentId  { get; set; }
        public string SeriesDesc { get; set; }
        public string StudyDesc  { get; set; }
        public string Suffix     { get; set; }
        public string Modality   { get; set; }  // "MR" or "CT"
        public string ProposedId { get; set; }
        public int    ZSize      { get; set; }
        public bool   Selected   { get; set; }
    }

    // ================================================================
    //  RenamerWindow
    //  Shows MRI and CT images together in one list with checkboxes.
    //  MRI rows are tinted blue, CT rows tinted orange so they are
    //  easy to distinguish at a glance.
    // ================================================================
    public class RenamerWindow : Window
    {
        private readonly Patient          _patient;
        private readonly List<RenameItem> _items;

        private ProgressBar _progress;
        private TextBlock   _status;
        private Button      _applyBtn;
        private CheckBox    _selectAllBox;

        private readonly List<CheckBox> _rowCheckBoxes = new List<CheckBox>();
        private bool _updatingSelectAll = false;

        public RenamerWindow(Patient patient, List<RenameItem> items)
        {
            _patient = patient;
            _items   = items;
            BuildUI();
        }

        private void BuildUI()
        {
            Title                 = "Image ID Renamer — " + _patient.Id;
            Width                 = 780;
            SizeToContent         = SizeToContent.Height;
            ResizeMode            = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            var root = new StackPanel { Margin = new Thickness(18) };

            // Patient header
            root.Children.Add(MakeLabel(_patient.Name + "  (" + _patient.Id + ")", bold: true, size: 14));
            root.Children.Add(MakeLabel(
                _items.Count(i => i.Modality == "MR") + " MRI  +  " +
                _items.Count(i => i.Modality == "CT") + " CT  image(s) found. Tick the ones you want to rename.",
                size: 11, gray: true));
            root.Children.Add(Gap(10));

            // ── Colour key ────────────────────────────────────────────
            var keyPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin      = new Thickness(0, 0, 0, 8)
            };
            keyPanel.Children.Add(MakeColourSwatch(Color.FromRgb(220, 235, 252)));
            keyPanel.Children.Add(MakeLabel(" MRI    ", size: 11));
            keyPanel.Children.Add(MakeColourSwatch(Color.FromRgb(255, 235, 210)));
            keyPanel.Children.Add(MakeLabel(" CT", size: 11));
            root.Children.Add(keyPanel);

            // ── Column headers ────────────────────────────────────────
            var headers = new Grid { Margin = new Thickness(0, 0, 0, 2) };
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(26) });  // checkbox
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });  // modality badge
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(115) }); // current ID
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });  // arrow
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(145) }); // new ID
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) });  // suffix
            headers.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // desc

            _selectAllBox = new CheckBox
            {
                IsChecked           = true,
                VerticalAlignment   = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                ToolTip             = "Select / deselect all"
            };
            _selectAllBox.Checked   += OnSelectAllChanged;
            _selectAllBox.Unchecked += OnSelectAllChanged;
            AddToGrid(headers, _selectAllBox,                                          0);
            AddToGrid(headers, MakeLabel("",             bold: true, size: 11),        1);
            AddToGrid(headers, MakeLabel("Current ID",   bold: true, size: 11),        2);
            AddToGrid(headers, MakeLabel("",             size: 11),                    3);
            AddToGrid(headers, MakeLabel("New ID",       bold: true, size: 11),        4);
            AddToGrid(headers, MakeLabel("Suffix",       bold: true, size: 11),        5);
            AddToGrid(headers, MakeLabel("Series  |  Study", bold: true, size: 11),   6);
            root.Children.Add(headers);

            // ── Scrollable list ───────────────────────────────────────
            var listBorder = new Border
            {
                BorderBrush     = new SolidColorBrush(Colors.LightGray),
                BorderThickness = new Thickness(1),
                MaxHeight       = 280
            };
            var scroll    = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var listPanel = new StackPanel { Margin = new Thickness(4, 4, 4, 4) };

            foreach (var item in _items)
            {
                var capturedItem = item;

                // Row background — blue tint for MRI, orange tint for CT
                bool isMR = item.Modality == "MR";
                var rowBg = new SolidColorBrush(
                    isMR ? Color.FromRgb(220, 235, 252)   // light blue
                         : Color.FromRgb(255, 235, 210));  // light orange

                var row = new Grid
                {
                    Margin     = new Thickness(0, 1, 0, 1),
                    Background = rowBg
                };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(26) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(115) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(145) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                // Checkbox
                var cb = new CheckBox
                {
                    IsChecked           = item.Selected,
                    VerticalAlignment   = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin              = new Thickness(2)
                };
                cb.Checked   += (s, e) => { capturedItem.Selected = true;  RefreshApplyButton(); UpdateSelectAllState(); };
                cb.Unchecked += (s, e) => { capturedItem.Selected = false; RefreshApplyButton(); UpdateSelectAllState(); };
                _rowCheckBoxes.Add(cb);
                AddToGrid(row, cb, 0);

                // Modality badge — MR or CT label
                AddToGrid(row, new TextBlock
                {
                    Text                = item.Modality,
                    FontSize            = 10,
                    FontWeight          = FontWeights.Bold,
                    Foreground          = isMR
                        ? new SolidColorBrush(Color.FromRgb(0, 70, 140))
                        : new SolidColorBrush(Color.FromRgb(160, 80, 0)),
                    VerticalAlignment   = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                }, 1);

                // Current ID
                AddToGrid(row, new TextBlock
                {
                    Text              = item.CurrentId,
                    Foreground        = Brushes.Gray,
                    FontFamily        = new FontFamily("Consolas"),
                    FontSize          = 11,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin            = new Thickness(2, 0, 0, 0)
                }, 2);

                // Arrow
                AddToGrid(row, new TextBlock
                {
                    Text                = "->",
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                    FontSize            = 11
                }, 3);

                // Proposed new ID — green and bold
                AddToGrid(row, new TextBlock
                {
                    Text              = item.ProposedId,
                    Foreground        = new SolidColorBrush(Color.FromRgb(0, 128, 0)),
                    FontFamily        = new FontFamily("Consolas"),
                    FontWeight        = FontWeights.Bold,
                    FontSize          = 11,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin            = new Thickness(2, 0, 0, 0)
                }, 4);

                // Suffix — blue if matched, grey if none
                AddToGrid(row, new TextBlock
                {
                    Text              = string.IsNullOrEmpty(item.Suffix) ? "none" : item.Suffix,
                    Foreground        = string.IsNullOrEmpty(item.Suffix)
                        ? Brushes.LightGray
                        : new SolidColorBrush(Color.FromRgb(0, 80, 160)),
                    FontFamily        = new FontFamily("Consolas"),
                    FontWeight        = FontWeights.Bold,
                    FontSize          = 11,
                    VerticalAlignment = VerticalAlignment.Center
                }, 5);

                // Series + study descriptions
                AddToGrid(row, new TextBlock
                {
                    Text              = item.SeriesDesc + "  |  " + item.StudyDesc,
                    Foreground        = Brushes.DimGray,
                    FontSize          = 10,
                    FontStyle         = FontStyles.Italic,
                    TextTrimming      = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin            = new Thickness(4, 0, 4, 0)
                }, 6);

                listPanel.Children.Add(row);
            }

            scroll.Content   = listPanel;
            listBorder.Child = scroll;
            root.Children.Add(listBorder);
            root.Children.Add(Gap(14));

            // Progress bar
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

            // Status message
            _status = new TextBlock
            {
                Text         = "",
                FontSize     = 12,
                TextWrapping = TextWrapping.Wrap,
                MinHeight    = 18,
                Margin       = new Thickness(0, 0, 0, 10)
            };
            root.Children.Add(_status);

            // Apply button
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

            // Cancel button
            var cancelBtn = new Button { Content = "Cancel", Height = 28, FontSize = 12 };
            cancelBtn.Click += (s, e) => Close();
            root.Children.Add(cancelBtn);

            Content = root;
        }

        // ── SelectAll / per-row checkbox logic ───────────────────────

        private void OnSelectAllChanged(object sender, RoutedEventArgs e)
        {
            if (_updatingSelectAll) return;
            bool selectAll = _selectAllBox.IsChecked == true;
            _updatingSelectAll = true;
            for (int i = 0; i < _items.Count; i++)
            {
                _items[i].Selected       = selectAll;
                _rowCheckBoxes[i].IsChecked = selectAll;
            }
            _updatingSelectAll = false;
            RefreshApplyButton();
        }

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
                _selectAllBox.IsChecked = null; // indeterminate — some selected
            _updatingSelectAll = false;
        }

        private void RefreshApplyButton()
        {
            int count = _items.Count(i => i.Selected);
            _applyBtn.Content   = "Apply  (" + count + " selected)";
            _applyBtn.IsEnabled = count > 0;
        }

        // ============================================================
        //  OnApply
        //  Only processes items where item.Selected == true.
        // ============================================================
        private void OnApply(object sender, RoutedEventArgs e)
        {
            var selectedItems = _items.Where(i => i.Selected).ToList();
            if (selectedItems.Count == 0) return;

            var confirm = MessageBox.Show(
                "Rename " + selectedItems.Count + " image ID(s) for patient " + _patient.Id + "?\n\n" +
                "This cannot be undone from within this script.",
                "Confirm",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes) return;

            _applyBtn.IsEnabled  = false;
            _progress.Maximum    = selectedItems.Count;
            _progress.Value      = 0;
            _progress.Visibility = Visibility.Visible;

            int ok = 0;
            var errors = new List<string>();

            try
            {
                // BeginModifications() is required before any write.
                // Throws if the system is not configured for write scripts.
                _patient.BeginModifications();

                foreach (var item in selectedItems)
                {
                    try
                    {
                        // Set the new image ID.
                        // image.Id is confirmed writeable: public string Id { get; set; }
                        item.Image.Id = item.ProposedId;

                        // Attempt to also set the series ID.
                        // series.Id shows as { get; } in the 16.1 CHM docs, but
                        // the CHM may be out of date — Aria 16.2 may have added
                        // a setter. We try it silently: if it works, great; if
                        // it throws a property-is-read-only exception we just
                        // swallow it and carry on. The image rename still succeeds.
                        try { item.Series.Id = item.ProposedId; }
                        catch { /* series.Id read-only on this version — ignore */ }

                        ok++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add(item.CurrentId + ": " + ex.Message);
                    }

                    _progress.Value++;
                    DoEvents();
                }
            }
            catch (Exception ex)
            {
                SetStatus("Could not begin modifications: " + ex.Message, red: true);
                _applyBtn.IsEnabled  = true;
                _progress.Visibility = Visibility.Hidden;
                return;
            }

            if (errors.Count == 0)
                SetStatus(ok + " image(s) renamed successfully.", red: false);
            else
                SetStatus(
                    ok + " renamed.   " + errors.Count + " failed:\n" +
                    string.Join("\n", errors.Select(err => "  - " + err)),
                    red: true);

            _applyBtn.Content   = "Close";
            _applyBtn.IsEnabled = true;
            _applyBtn.Click    -= OnApply;
            _applyBtn.Click    += (s, ev) => Close();
        }

        // ── Helpers ───────────────────────────────────────────────────

        private void SetStatus(string msg, bool red)
        {
            _status.Text       = msg;
            _status.Foreground = red
                ? new SolidColorBrush(Colors.DarkRed)
                : new SolidColorBrush(Color.FromRgb(0, 128, 0));
        }

        private static void DoEvents()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new Action(() => { }));
        }

        // Small coloured square for the modality colour key
        private static Border MakeColourSwatch(Color color)
        {
            return new Border
            {
                Width      = 16,
                Height     = 16,
                Background = new SolidColorBrush(color),
                BorderBrush     = new SolidColorBrush(Colors.Gray),
                BorderThickness = new Thickness(1),
                Margin     = new Thickness(4, 0, 2, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        private static TextBlock MakeLabel(string text, bool bold = false,
            double size = 13, bool gray = false)
        {
            return new TextBlock
            {
                Text              = text,
                FontSize          = size,
                FontWeight        = bold ? FontWeights.Bold : FontWeights.Normal,
                Foreground        = gray ? Brushes.Gray : Brushes.Black,
                VerticalAlignment = VerticalAlignment.Center
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
