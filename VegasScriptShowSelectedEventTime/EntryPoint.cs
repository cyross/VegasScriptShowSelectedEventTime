using ScriptPortal.Vegas;
using System.Windows.Forms;
using VegasScriptHelper;

namespace VegasScriptShowSelectedEventTime
{
    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            VegasScriptSettings.Load();
            VegasHelper helper = VegasHelper.Instance(vegas);

            try
            {
                TrackEvent e = helper.GetSelectedEvent();
                long start_nanos = helper.GetEventStartTime(e) + 500000 / 1000000 * 1000000;
                long length_nanos = helper.GetEventLength(e) + 500000 / 1000000 * 1000000;
                MessageBox.Show(
                    string.Format(
                        "開始:{0} 長さ:{1}",
                        VegasHelperUtility.NanoToTimestamp(start_nanos),
                        VegasHelperUtility.NanoToTimestamp(length_nanos)
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
