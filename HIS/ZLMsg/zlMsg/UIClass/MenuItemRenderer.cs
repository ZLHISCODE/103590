using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace zlShortMsg
{
    public class MenuItemRenderer : ToolStripRenderer
    {
        public MenuItemRenderer() : base()
        {

        }

        //重写Render修改背景色方法
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected == true && e.Item.Enabled)
            {
                DrawMenuDropDownItemHighlight(e);
                e.Item.ForeColor = Color.Black;
            }
            else
            {
                base.OnRenderMenuItemBackground(e);
                if (e.Item.IsOnDropDown)
                    e.Item.ForeColor = SystemColors.ControlText;
                else
                    e.Item.ForeColor = SystemColors.ControlLightLight;
            }
        }

        private void DrawMenuDropDownItemHighlight(ToolStripItemRenderEventArgs e)
        {
            Rectangle rect = new Rectangle();
            rect = new Rectangle(2, 0, (int)e.Graphics.VisibleClipBounds.Width - 4, (int)e.Graphics.VisibleClipBounds.Height - 1);
            using (LinearGradientBrush b = new LinearGradientBrush(rect, ColorTranslator.FromHtml("#FFFF66"), ColorTranslator.FromHtml("#FFFFCC"), (float)10))
            {
                e.Graphics.FillRectangle(b, rect);
            }
        }

    }
}