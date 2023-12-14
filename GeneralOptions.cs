using Microsoft.VisualStudio.Shell;
using System.ComponentModel;

namespace SSMSExec
{
    internal class GeneralOptions : DialogPage
    {
        [Category("General")]
        [DisplayName("Exe Command 1 Active?")]
        [Description("Hide or show the button")]
        public bool ExeActive { get; set; } = true;

        [Category("General")]
        [DisplayName("Button Text")]
        [Description("The name of the command to display in the menu")]
        public string ButtonText { get; set; } = "Update Query";

        [Category("General")]
        [DisplayName("Exe Location")]
        [Description("The location of the exe which will update the current query")]
        public string ExeLocation { get; set; } = "";

        [Category("General")]
        [DisplayName("Exe Parameter 1")]
        [Description("A single parameter to supply when calling exe")]
        public string ExeParameter1 { get; set; } = "";

        [Category("General")]
        [DisplayName("Exe Parameter 2")]
        [Description("A single parameter to supply when calling exe")]
        public string ExeParameter2 { get; set; } = "";

        [Category("General")]
        [DisplayName("Exe Parameter 3")]
        [Description("A single parameter to supply when calling exe")]
        public string ExeParameter3 { get; set; } = "";

        [Category("General")]
        [DisplayName("Exe Parameter 4")]
        [Description("A single parameter to supply when calling exe")]
        public string ExeParameter4 { get; set; } = "";

        [Category("General")]
        [DisplayName("Exe Parameter 5")]
        [Description("A single parameter to supply when calling exe")]
        public string ExeParameter5 { get; set; } = "";

        [Category("General")]
        [DisplayName("Exe Parameter 6")]
        [Description("A single parameter to supply when calling exe")]
        public string ExeParameter6 { get; set; } = "";
    }
}
