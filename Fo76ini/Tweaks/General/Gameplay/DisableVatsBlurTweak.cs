using System;

namespace Fo76ini.Tweaks.General.Gameplay
{
    class DisableVatsBlurTweak : ITweak<bool>, ITweakInfo
    {
        public string Description => "Disables the blur effect when entering VATS.";
        public WarnLevel WarnLevel => WarnLevel.None;
        public string AffectedFiles => "Fallout76Custom.ini";
        public string AffectedValues => "[VATS]bVATSBlur";
        public bool DefaultValue => false;
        public string Identifier => this.GetType().FullName;
        public bool UIReloadNecessary => false;

        public bool GetValue()
        {
            // Returns true if VATS blur is disabled (bVATSBlur = 0)
            return IniFiles.GetInt("VATS", "bVATSBlur", 1) == 0;
        }

        public void SetValue(bool value)
        {
            if (value)
            {
                // Disable blur
                IniFiles.F76Custom.Set("VATS", "bVATSBlur", 0);
            }
            else
            {
                // Enable blur / revert to default
                IniFiles.F76Custom.Set("VATS", "bVATSBlur", 1);
            }
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
