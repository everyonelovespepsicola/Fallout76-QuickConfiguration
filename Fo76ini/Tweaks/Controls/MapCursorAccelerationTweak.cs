using System;

namespace Fo76ini.Tweaks.Controls
{
    class MapCursorAccelerationTweak : ITweak<float>, ITweakInfo
    {
        public string Description => "Changes the cursor acceleration speed when moving around the map.";
        public WarnLevel WarnLevel => WarnLevel.None;
        public string AffectedFiles => "Fallout76Custom.ini";
        public string AffectedValues => "[Controls]fMapMoveAcceleration";
        public float DefaultValue => 27.0f;
        public string Identifier => this.GetType().FullName;
        public bool UIReloadNecessary => false;

        public float GetValue()
        {
            return IniFiles.GetFloat("Controls", "fMapMoveAcceleration", DefaultValue);
        }

        public void SetValue(float value)
        {
            IniFiles.F76Custom.Set("Controls", "fMapMoveAcceleration", value);
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
