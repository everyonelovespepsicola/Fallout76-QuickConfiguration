using System;

namespace Fo76ini.Tweaks.Video
{
    class SunShadowUpdateTimeTweak : ITweak<bool>, ITweakInfo
    {
        public string Description => "Slows down sun shadow updates (sets fSunShadowConstantUpdateTime to 1.0 and fSunShadowLinearUpdateTime to 1.0) under [Display]. This smooths out shadow movements, stopping shadows from jumping and stuttering across the screen.";

        public WarnLevel WarnLevel => WarnLevel.Notice;

        public string AffectedFiles => "Fallout76Custom.ini";

        public string AffectedValues => "[Display]fSunShadowConstantUpdateTime, [Display]fSunShadowLinearUpdateTime";

        public bool DefaultValue => false;

        public string Identifier => this.GetType().FullName;

        public bool UIReloadNecessary => false;

        public bool GetValue()
        {
            float constantTime = IniFiles.GetFloat("Display", "fSunShadowConstantUpdateTime", 0.1f);
            return Math.Abs(constantTime - 1.0f) < 0.05f;
        }

        public void SetValue(bool value)
        {
            if (value)
            {
                IniFiles.F76Custom.Set("Display", "fSunShadowConstantUpdateTime", 1.0f);
                IniFiles.F76Custom.Set("Display", "fSunShadowLinearUpdateTime", 1.0f);
            }
            else
            {
                IniFiles.F76Custom.Remove("Display", "fSunShadowConstantUpdateTime");
                IniFiles.F76Custom.Remove("Display", "fSunShadowLinearUpdateTime");
            }
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
