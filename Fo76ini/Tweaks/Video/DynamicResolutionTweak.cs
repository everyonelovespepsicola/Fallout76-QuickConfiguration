using System;

namespace Fo76ini.Tweaks.Video
{
    class DynamicResolutionTweak : ITweak<bool>, ITweakInfo
    {
        public string Description => "Enables bDynamicResolutionEnabled under [Display]. This allows dynamic resolution scaling to boost frame rates when the GPU is under heavy load.";

        public WarnLevel WarnLevel => WarnLevel.Notice;

        public string AffectedFiles => "Fallout76Custom.ini";

        public string AffectedValues => "[Display]bDynamicResolutionEnabled";

        public bool DefaultValue => false;

        public string Identifier => this.GetType().FullName;

        public bool UIReloadNecessary => false;

        public bool GetValue()
        {
            return IniFiles.GetBool("Display", "bDynamicResolutionEnabled", DefaultValue);
        }

        public void SetValue(bool value)
        {
            IniFiles.F76Custom.Set("Display", "bDynamicResolutionEnabled", value);
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
