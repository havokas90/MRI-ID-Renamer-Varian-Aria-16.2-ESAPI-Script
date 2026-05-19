using System;
using System.Collections.Generic;
using System.Windows;
using VMS.TPS;

namespace ImageIdRenamer.LocalPreview
{
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            var app = new Application();
            app.Run(new ImageIdRenamerWindow(BuildSampleCandidates()));
        }

        private static List<RenameCandidate> BuildSampleCandidates()
        {
            return new List<RenameCandidate>
            {
                Candidate("MR", "MR1", "T2_AX_HN", "_HN", "t2_ax_2.0mm", "Head and Neck MRI"),
                Candidate("MR", "MR2", "DWI_HN", "_HN", "ep2d_diff_b1000_3.0mm", "Head and Neck MRI"),
                Candidate("MR", "MR3", "T1_POST_HN", "_HN", "t1 post contrast", "Head and Neck MRI"),
                Candidate("CT", "CT1", "CT_SIM_PEL", "_PEL", "CT SIM 2.5mm", "Pelvis planning CT"),
                Candidate("CT", "CT2", "CT_CHEST_CH", "_CH", "4D CT chest", "Thorax 4DCT"),
                Candidate("MR", "OLD_LONG_NAME", "T2_SPACE_BRAIN_BR", "_BR", "T2 SPACE brain 1.0mm", "Brain MRI")
            };
        }

        private static RenameCandidate Candidate(string modality, string currentId, string newId, string suffix, string series, string study)
        {
            return new RenameCandidate
            {
                Modality = modality,
                CurrentId = currentId,
                NewId = newId,
                Suffix = suffix,
                SeriesDescription = series,
                StudyDescription = study,
                IsSelected = true
            };
        }
    }
}
