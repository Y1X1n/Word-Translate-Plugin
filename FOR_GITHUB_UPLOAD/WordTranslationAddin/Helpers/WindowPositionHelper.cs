using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using WordTranslationAddin.Models;

namespace WordTranslationAddin.Helpers
{
    public static class WindowPositionHelper
    {
        public static Point CalculatePopupPosition(
            PopupPosition position, 
            Point cursorPosition, 
            int popupWidth, 
            int popupHeight,
            System.Drawing.Rectangle screenBounds)
        {
            int x = cursorPosition.X;
            int y = cursorPosition.Y;

            switch (position)
            {
                case PopupPosition.Auto:
                    // 自动定位：优先显示在光标右下方，如果空间不足则调整
                    x = cursorPosition.X + 10;
                    y = cursorPosition.Y + 10;
                    
                    // 检查右边界
                    if (x + popupWidth > screenBounds.Right)
                    {
                        x = cursorPosition.X - popupWidth - 10;
                    }
                    
                    // 检查下边界
                    if (y + popupHeight > screenBounds.Bottom)
                    {
                        y = cursorPosition.Y - popupHeight - 10;
                    }
                    break;

                case PopupPosition.Cursor:
                    x = cursorPosition.X;
                    y = cursorPosition.Y;
                    break;

                case PopupPosition.Center:
                    x = screenBounds.Left + (screenBounds.Width - popupWidth) / 2;
                    y = screenBounds.Top + (screenBounds.Height - popupHeight) / 2;
                    break;

                case PopupPosition.TopLeft:
                    x = screenBounds.Left + 20;
                    y = screenBounds.Top + 20;
                    break;

                case PopupPosition.TopRight:
                    x = screenBounds.Right - popupWidth - 20;
                    y = screenBounds.Top + 20;
                    break;

                case PopupPosition.BottomLeft:
                    x = screenBounds.Left + 20;
                    y = screenBounds.Bottom - popupHeight - 20;
                    break;

                case PopupPosition.BottomRight:
                    x = screenBounds.Right - popupWidth - 20;
                    y = screenBounds.Bottom - popupHeight - 20;
                    break;
            }

            // 确保不超出屏幕边界
            x = Math.Max(screenBounds.Left, Math.Min(x, screenBounds.Right - popupWidth));
            y = Math.Max(screenBounds.Top, Math.Min(y, screenBounds.Bottom - popupHeight));

            return new Point(x, y);
        }

        public static System.Drawing.Rectangle GetCurrentScreenBounds(Point point)
        {
            var screen = Screen.FromPoint(point);
            return screen.WorkingArea;
        }

        public static void EnsureWindowVisibility(ref int left, ref int top, int width, int height)
        {
            var virtualScreen = SystemInformation.VirtualScreen;
            
            left = Math.Max(virtualScreen.Left, Math.Min(left, virtualScreen.Right - width));
            top = Math.Max(virtualScreen.Top, Math.Min(top, virtualScreen.Bottom - height));
        }
    }
}
