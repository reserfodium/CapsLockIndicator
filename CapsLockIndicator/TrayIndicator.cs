﻿using System;
using System.Drawing;
using System.Windows.Forms;

namespace CapsLockIndicator
{
    class TrayIndicator : IDisposable
    {
        #region Data

        NotifyIcon trayIcon;
        Timer timer;

        #region Font

        string IconFontName { get; } = "Pixel FJVerdana";
        int IconFontSize { get; } = 14;
        FontStyle IconFontStyle { get; } = FontStyle.Regular;

        #endregion

        #region Common

        Icon capsEnabledIcon, capsDisabledIcon;

        bool CapsLockState => Control.IsKeyLocked(Keys.CapsLock);
        bool CachedCapsLockState = false;

        Color MainColor { get; } = Color.White;

        float OffsetX { get; } = 2;
        float OffsetY { get; } = -1;

        int MagicSize { get; } = 16;  // Constant tray icon size 

        #endregion

        #endregion
        
        public TrayIndicator()
        {
            // Create icons
            capsEnabledIcon = GenerateIcon("A");
            capsDisabledIcon = GenerateIcon("a");

            trayIcon = new NotifyIcon();
            trayIcon.ContextMenuStrip = CreateContextMenu();
            SetTrayIcon();

            timer = new Timer();
            timer.Enabled = false;
            timer.Tick += timer_Update;
        }

        private void timer_Update(object sender, EventArgs e)
        {
            if (CapsLockState != CachedCapsLockState)
            {
                SetTrayIcon();
                CachedCapsLockState = CapsLockState;
            }
        }

        #region Functions

        public void Display()
        {
            trayIcon.Visible = true;
            timer.Enabled = true;
        }

        private ContextMenuStrip CreateContextMenu()
        {
            var menu = new ContextMenuStrip();

            // Exit
            ToolStripMenuItem exit = new ToolStripMenuItem();
            exit.Text = "Exit";
            exit.Click += (sender, e) => Application.Exit();
            menu.Items.Add(exit);

            return menu;
        }

        private void SetTrayIcon() => trayIcon.Icon = CapsLockState ? capsEnabledIcon : capsDisabledIcon;

        private Icon GenerateIcon(string text)
        {
            Font fontToUse = new Font(IconFontName, IconFontSize, IconFontStyle, GraphicsUnit.Pixel);
            Brush brushToUse = new SolidBrush(MainColor);
            Bitmap bitmapText = new Bitmap(MagicSize, MagicSize);  // Const size for tray icon

            Graphics g = Graphics.FromImage(bitmapText);

            g.Clear(Color.Transparent);

            // Draw border
            g.DrawLine(
                new Pen(MainColor, 1),
                0, 15, 15, 15);

            // Draw text
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit;
            g.DrawString(text, fontToUse, brushToUse, OffsetX, OffsetY);

            // Create icon from bitmap and return it
            // bitmapText.GetHicon() can throw exception
            return Icon.FromHandle(bitmapText.GetHicon());
        }

        #endregion

        public void Dispose()
        {
            trayIcon.Dispose();
            timer.Dispose();
        }
    }
}