using ScriptPortal.Vegas;
using System;
using System.Collections;
using System.Windows.Forms;
using VegasScriptHelper;

namespace VegasScriptShowSelectedEventTime
{
    public class CustomModule : ICustomCommandModule
    {
        private VegasHelper helper;
        private readonly static string CommandName = "ShowTrackLength";
        private readonly static string DisplayName = "選択したイベントの開始位置・長さを表示";
        private readonly static string DockName = "イベントの開始位置・長さ";
        private CustomCommand myCommand = new CustomCommand(CommandCategory.Tools, CommandName);

        public void InitializeModule(Vegas vegas)
        {
            helper = VegasHelper.Instance(vegas);
        }

        public ICollection GetCustomCommands()
        {
            helper.AddTrackEventStateChangedEventHandler(ShowDockView); // イベントをクリックすると自動的に表示される
            myCommand.DisplayName = DisplayName;
            myCommand.Invoked += HandleInvoked;
            myCommand.MenuPopup += HandleMenuPopup;
            return new CustomCommand[] { myCommand };
        }

        void HandleInvoked(Object sender, EventArgs e)
        {
            ShowDockView(sender, e);
        }

        void ShowDockView(Object sender, EventArgs e)
        {
            string result1 = "";
            string result2 = "";
            try
            {
                TrackEvent ev = helper.GetSelectedEvent();
                result1 = VegasHelperUtility.NanoToTimestamp(VegasHelperUtility.RoundNanos(helper.GetEventStartTime(ev)));
                result2 = VegasHelperUtility.NanoToTimestamp(VegasHelperUtility.RoundNanos(helper.GetEventLength(ev)));
            }
            catch (VegasHelperTrackUnselectedException)
            {
                // 空文字列のままで良いのでpass
            }
            catch (VegasHelperNoneEventsException)
            {
                // 空文字列のままで良いのでpass
            }
            catch (VegasHelperNoneSelectedEventException)
            {
                // 空文字列のままで良いのでpass
            }
            if (!helper.ActivateDockView(DockName))
            {
                LoadDockView(result1, result2);
            }
            else
            {
                UpdateDockView(result1, result2);
            }
        }
        void LoadDockView(string result1, string result2)
        {
            DockableControl dock = new DockableControl(DockName);

            FlowLayoutPanel panel = new FlowLayoutPanel();
            panel.Dock = DockStyle.Fill;

            Label label1 = CreateLabel("Result1", GetStartTimeString(result1));
            panel.Controls.Add(label1);

            Label label2 = CreateLabel("Result2", GetLengthString(result2));
            label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            panel.Controls.Add(label2);

            dock.Controls.Add(panel);

            helper.LoadDockView(dock);
        }

        void HandleMenuPopup(Object sender, EventArgs e)
        {
            myCommand.Checked = helper.FindDockView(DockName);
        }

        void UpdateDockView(string result1, string result2)
        {
            IDockView dockView = null;
            if(!helper.FindDockView(DockName, ref dockView))
            {
                LoadDockView(result1, result2);
                return;
            }
            DockableControl dock = (DockableControl)dockView;
            ((Label)(dock.Controls[0].Controls[0])).Text = GetStartTimeString(result1);
            ((Label)(dock.Controls[0].Controls[1])).Text = GetLengthString(result2);
        }

        private Label CreateLabel(string name, string text)
        {
            Label label = new Label();
            label.Name = name;
            label.Dock = DockStyle.Fill;
            label.AutoSize = true;
            label.Text = text;
            label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            return label;
        }

        private string GetStartTimeString(string timeString)
        {
            // ウインドウをドッキングさせるとタイトルが隠れるため「イベントの」で明示
            return string.Format("イベントの開始時間:{0}", timeString);
        }

        private string GetLengthString(string timeString)
        {
            // ウインドウをドッキングさせるとタイトルが隠れるため「イベントの」で明示
            return string.Format("イベントの長さ:{0}", timeString);
        }
    }
}
