using System;

namespace Fo76ini.Tweaks.Controls
{
    class MapCursorTweaksActiveTweak : ITweak<bool>, ITweakInfo
    {
        private readonly MapCursorAccelerationTweak accTweak = new MapCursorAccelerationTweak();
        private readonly MapCursorSpeedTweak speedTweak = new MapCursorSpeedTweak();

        public string Description => "Optimizes map mouse cursor movement by reducing speed and acceleration (sets fMapMoveAcceleration to 7 and fMapMoveSpeed to 5).";
        public WarnLevel WarnLevel => WarnLevel.None;
        public string AffectedFiles => "Fallout76Custom.ini";
        public string AffectedValues => "[Controls]fMapMoveAcceleration, [Controls]fMapMoveSpeed";
        public bool DefaultValue => false;
        public string Identifier => this.GetType().FullName;
        public bool UIReloadNecessary => false;

        public bool GetValue()
        {
            float acceleration = accTweak.GetValue();
            float speed = speedTweak.GetValue();
            // True if they are set to custom/improved values
            return Math.Abs(acceleration - 7.0f) < 0.1f && Math.Abs(speed - 5.0f) < 0.1f;
        }

        public void SetValue(bool value)
        {
            if (value)
            {
                accTweak.SetValue(7.0f);
                speedTweak.SetValue(5.0f);
            }
            else
            {
                // Revert to game defaults
                accTweak.SetValue(27.0f);
                speedTweak.SetValue(10.0f);
            }
        }

        public void ResetValue()
        {
            SetValue(DefaultValue);
        }
    }
}
