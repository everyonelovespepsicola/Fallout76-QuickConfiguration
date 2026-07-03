using System;

namespace Fo76ini.Tweaks.Controls
{
    class MapCursorSpeedTweak : ITweak<float>, ITweakInfo
    {
        public string Description => "Changes the cursor speed when moving around the map.";
        public WarnLevel WarnLevel => WarnLevel.None;
        public string AffectedFiles => "Fallout76Custom.ini";
        public string AffectedValues => "[Client]fMapCursorMoveSpeed";
        public float DefaultValue => 10.0f;
        public string Identifier => this.GetType().FullName;
        public bool UIReloadNecessary => false;

        public float GetValue()
        {
            return IniFiles.GetFloat("Client", "fMapCursorMoveSpeed", DefaultValue);
        }

        public void SetValue(float value)
        {
            IniFiles.F76Custom.Set("Client", "fMapCursorMoveSpeed", value);
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
