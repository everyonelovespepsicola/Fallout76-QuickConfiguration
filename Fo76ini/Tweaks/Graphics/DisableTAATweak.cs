using System;

namespace Fo76ini.Tweaks.Graphics
{
    class DisableTAATweak : ITweak<bool>, ITweakInfo
    {
        private readonly AntiAliasingTweak aaTweak = new AntiAliasingTweak();

        public string Description => "Completely disables Temporal Anti-Aliasing (TAA) to remove image blur/smearing when moving the camera.";
        public WarnLevel WarnLevel => WarnLevel.None;
        public string AffectedFiles => "Fallout76Prefs.ini";
        public string AffectedValues => "[Display]sAntiAliasing";
        public bool DefaultValue => false;
        public string Identifier => this.GetType().FullName;
        public bool UIReloadNecessary => false;

        public bool GetValue()
        {
            return aaTweak.GetValue() == AntiAliasing.Disabled;
        }

        public void SetValue(bool value)
        {
            if (value)
            {
                aaTweak.SetValue(AntiAliasing.Disabled);
            }
            else
            {
                aaTweak.SetValue(AntiAliasing.TAA);
            }
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
