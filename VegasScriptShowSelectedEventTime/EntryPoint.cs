using ScriptPortal.Vegas;
using System.Windows.Forms;
using VegasScriptHelper;

namespace VegasScriptShowSelectedEventTime
{
    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            VegasHelper helper = VegasHelper.Instance(vegas);

            try
            {
                TrackEvent e = helper.GetSelectedEvent();
                Timecode event_start_time = helper.GetEventStartTime(e);
                Timecode event_length = helper.GetEventLength(e);
                MessageBox.Show(
                    string.Format(
                        "開始:{0} 長さ:{1}", 
                        event_start_time.ToString(),
                        event_length.ToString()
                        )
                    );
            }
            catch (VegasHelperTrackUnselectedException)
            {
                MessageBox.Show("トラックが選択されていません。");
            }
            catch (VegasHelperNoneEventsException)
            {
                MessageBox.Show("選択したトラック中にイベントが存在していません。");
            }
            catch(VegasHelperNoneSelectedEventException)
            {
                MessageBox.Show("イベントが選択されていません。");
            }
        }
    }
}
