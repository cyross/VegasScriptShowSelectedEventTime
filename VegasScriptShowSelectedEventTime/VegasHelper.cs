using ScriptPortal.Vegas;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Windows.Forms;

namespace VegasScriptShowSelectedEventTime
{
    internal struct VegasDuration
    {
        public Timecode StartTime;
        public Timecode Length;
    }

    /// <summary>
    /// Vegasオブジェクトを操作するヘルパクラス
    /// 本クラスはSingleton
    /// </summary>
    internal partial class VegasHelper
    {
        private static VegasHelper _instance = null;

        internal Vegas Vegas { get; set; }

        internal readonly Timecode BaseTimecode = new Timecode();

        internal static VegasHelper Instance(Vegas vegas)
        {
            if (_instance == null)
            {
                _instance = new VegasHelper(vegas);
            }
            else
            {
                _instance.Vegas = vegas;
            }

            return _instance;
        }

        private VegasHelper(Vegas vegas)
        {
            Vegas = vegas;
        }

        /// <summary>
        /// 現在VEGASが開いているプロジェクトを取得する
        /// </summary>
        internal Project Project
        {
            get
            {
                return Vegas.Project;
            }
        }

        internal bool ActivateDockView(string dockName)
        {
            return Vegas.ActivateDockView(dockName);
        }

        internal void LoadDockView(DockableControl dockView)
        {
            Vegas.LoadDockView(dockView);
        }

        internal bool FindDockView(string dockName)
        {
            return Vegas.FindDockView(dockName);
        }

        internal bool FindDockView(string dockname, ref IDockView dockView)
        {
            return Vegas.FindDockView(dockname, out dockView);
        }

        /// <summary>
        /// プロジェクト内で選択しているトラックがあれば、そのトラックのオブジェクトを返す。
        /// なければnullを返す
        /// </summary>
        /// <returns>選択プロジェクトがあればそのTrackオブジェクト、なければnull</returns>
        internal Track SelectedTrack()
        {
            return SelectedTrack(Vegas.Project);
        }

        /// <summary>
        /// プロジェクト内で選択しているトラックがあれば、そのトラックのオブジェクトを返す。
        /// なければnullを返す
        /// </summary>
        /// <param name="project">VEGASが開いているプロジェクト</param>
        /// <returns>選択プロジェクトがあればそのTrackオブジェクト、なければnull</returns>
        internal Track SelectedTrack(Project project)
        {
            foreach (Track track in project.Tracks)
            {
                if (track.Selected)
                {
                    return track;
                }
            }
            throw new VegasHelperTrackUnselectedException();
        }

        /// <summary>
        /// プロジェクト内で選択しているトラックがあれば、そのトラックのオブジェクトを返す。
        /// なければnullを返す
        /// </summary>
        /// <returns>選択プロジェクトがあればそのTrackオブジェクト、なければnull</returns>
        internal VideoTrack SelectedVideoTrack()
        {
            return SelectedVideoTrack(Vegas.Project);
        }

        /// <summary>
        /// プロジェクト内で選択しているトラックがあれば、そのトラックのオブジェクトを返す。
        /// なければnullを返す
        /// </summary>
        /// <param name="project">VEGASが開いているプロジェクト</param>
        /// <returns>選択プロジェクトがあればそのTrackオブジェクト、なければnull</returns>
        internal VideoTrack SelectedVideoTrack(Project project)
        {
            foreach (Track track in project.Tracks)
            {
                if (track.Selected && track.IsVideo())
                {
                    return (VideoTrack)track;
                }
            }
            throw new VegasHelperTrackUnselectedException();
        }

        /// <summary>
        /// プロジェクト内で選択しているトラックがあれば、そのトラックのオブジェクトを返す。
        /// なければnullを返す
        /// </summary>
        /// <returns>選択プロジェクトがあればそのTrackオブジェクト、なければnull</returns>
        internal AudioTrack SelectedAudioTrack()
        {
            return SelectedAudioTrack(Vegas.Project);
        }

        /// <summary>
        /// プロジェクト内で選択しているトラックがあれば、そのトラックのオブジェクトを返す。
        /// なければnullを返す
        /// </summary>
        /// <param name="project">VEGASが開いているプロジェクト</param>
        /// <returns>選択プロジェクトがあればそのTrackオブジェクト、なければnull</returns>
        internal AudioTrack SelectedAudioTrack(Project project)
        {
            foreach (Track track in project.Tracks)
            {
                if (track.Selected && track.IsAudio())
                {
                    return (AudioTrack)track;
                }
            }
            throw new VegasHelperTrackUnselectedException();
        }

        internal string GetTrackTitle(Track track)
        {
            return track.Name;
        }

        internal string GetVideoTrackTitle()
        {
            VideoTrack track = SelectedVideoTrack();
            return GetTrackTitle(track);
        }

        internal string GetAudioTrackTitle()
        {
            AudioTrack track = SelectedAudioTrack();
            return GetTrackTitle(track);
        }

        internal void SetTrackTitle(Track track, string title)
        {
            track.Name = title;
        }

        internal void SetVideoTrackTitle(string title)
        {
            VideoTrack track = SelectedVideoTrack();
            SetTrackTitle(track, title);
        }

        internal void SetAudioTrackTitle(string title)
        {
            AudioTrack track = SelectedAudioTrack();
            SetTrackTitle(track, title);
        }

        internal VideoTrack SearchVideoTrackByName(string name)
        {
            Project project = Vegas.Project;
            foreach (Track track in project.Tracks)
            {
                if (track.IsVideo() && track.Name == name)
                {
                    return (VideoTrack)track;
                }
            }
            throw new VegasHelperNotFoundTrackException();
        }

        internal AudioTrack SearchAudioTrackByName(string name)
        {
            Project project = Vegas.Project;
            foreach (Track track in project.Tracks)
            {
                if (track.IsAudio() && track.Name == name)
                {
                    return (AudioTrack)track;
                }
            }
            throw new VegasHelperNotFoundTrackException();
        }

        /// <summary>
        /// 引数で指定したトラックがビデオトラックかどうかを調べる
        /// </summary>
        /// <param name="track">対象のトラックオブジェクト</param>
        /// <returns>ビデオトラックの場合はTrue、それ以外のときはFalseを返す</returns>
        internal bool IsVideoTrack(Track track)
        {
            return track.IsVideo();
        }

        /// <summary>
        /// 引数で指定したトラックがオーディオトラックかどうかを調べる
        /// </summary>
        /// <param name="track">対象のトラックオブジェクト</param>
        /// <returns>オーディオトラックの場合はTrue、それ以外のときはFalseを返す</returns>
        internal bool IsAudioTrack(Track track)
        {
            return track.IsAudio();
        }

        internal VideoTrack AddVideoTrack()
        {
            return Vegas.Project.AddVideoTrack();
        }

        internal AudioTrack AddAudioTrack()
        {
            return Vegas.Project.AddAudioTrack();
        }

        /// <summary>
        /// オーディオトラックを作り、指定したディレクトリ内のwavファイルをイベントとして挿入する
        /// オーディオファイルの検知は拡張子のみで、ファイルの中身はチェックしない
        /// 対応するファイルはVegasScriptSettings.SupportedAudioFileで指定されたもの
        /// </summary>
        /// <param name="fileDir">指定したディレクトリ名</param>
        /// <param name="interval">挿入するイベント間の間隔　単位はミリ秒　標準は0.0</param>
        /// <param name="fromStart">トラックの最初から挿入するかどうかを示すフラグ　trueのときは最初から、falseのときは現在のカーソル位置から</param>
        /// <param name="recursive">子ディレクトリのを再帰的にトラックの最初から挿入するかどうかを示すフラグ　trueのときは最初から、falseのときは現在のカーソル位置から</param>
        internal void InseretAudioInTrack(string fileDir, float interval = 0.0f, bool fromStart = false, bool recursive = true)
        {
            AudioTrack audioTrack = AddAudioTrack();
            SetTrackTitle(audioTrack, "Subtitles");
            audioTrack.Selected = true;

            Timecode currentPosition = fromStart ? new Timecode() : Vegas.Cursor;
            Timecode intervalTimecode = new Timecode(interval);

            _InsertAudio(currentPosition, intervalTimecode, fileDir, audioTrack, recursive);
        }

        private Timecode _InsertAudio(Timecode current, Timecode interval, string fileDir, AudioTrack audioTrack, bool recursive)
        {
            if (recursive)
            {
                foreach (string childDir in Directory.GetDirectories(fileDir))
                {
                    current = _InsertAudio(current, interval, childDir, audioTrack, recursive);
                }
            }
            foreach (string filePath in Directory.GetFiles(fileDir))
            {
                if (VegasScriptSettings.SupportedAudioFile.Contains(Path.GetExtension(filePath)))
                {
                    Media audioMedia = new Media(filePath);
                    AudioStream audioStream = audioMedia.GetAudioStreamByIndex(0);

                    AudioEvent audioEvent = audioTrack.AddAudioEvent(current, audioStream.Length);
                    audioEvent.AddTake(audioStream);

                    current += audioStream.Length + interval;
                }
            }
            return current;
        }

        internal TrackEvents GetEvents(Track track, bool throwError = true)
        {
            if (throwError && track.Events.Count == 0) { throw new VegasHelperNoneEventsException(); }
            return track.Events;
        }

        internal TrackEvents GetVideoEvents(bool throwError = true)
        {
            VideoTrack selected = SelectedVideoTrack();
            if (throwError && selected.Events.Count == 0) { throw new VegasHelperNoneEventsException(); }
            return selected.Events;
        }

        internal TrackEvents GetAudioEvents(bool throwError = true)
        {
            AudioTrack selected = SelectedAudioTrack();
            if (throwError && selected.Events.Count == 0) { throw new VegasHelperNoneEventsException(); }
            return selected.Events;
        }

        internal TrackEvent GetSelectedEvent(Track track, bool throwError = true)
        {
            if (throwError && track.Events.Count == 0) { throw new VegasHelperNoneEventsException(); }
            foreach (TrackEvent e in track.Events)
            {
                if (e.Selected)
                {
                    return e;
                }
            }
            if (throwError) { throw new VegasHelperNoneSelectedEventException(); }
            return null;
        }

        internal TrackEvent GetSelectedEvent(bool throwError = true)
        {
            Track track = SelectedTrack();
            return GetSelectedEvent(track, throwError);
        }

        internal TrackEvent GetSelectedVideoEvent(bool throwError = true)
        {
            Track track = SelectedVideoTrack();
            return GetSelectedEvent(track, throwError);
        }

        internal TrackEvent GetSelectedAudioEvent(bool throwError = true)
        {
            Track track = SelectedAudioTrack();
            return GetSelectedEvent(track, throwError);
        }

        internal long GetEventStartTime(TrackEvent trackEvent)
        {
            return trackEvent.Start.Nanos;
        }

        internal void SetEventStartTime(TrackEvent trackEvent, long nanos)
        {
            trackEvent.Start = new Timecode(nanos);
        }

        internal long GetEventLength(TrackEvent trackEvent)
        {
            return trackEvent.Length.Nanos;
        }

        internal void SetEventLength(TrackEvent trackEvent, long nanos)
        {
            trackEvent.Length = new Timecode(nanos);
        }

        internal Take[] GetFirstTakes(Track track)
        {
            return GetFirstTakes(track.Events);
        }

        internal Take[] GetFirstTakes(TrackEvents events)
        {
            IEnumerable<Take> takes = events.Select(e => GetFirstTake(e));
            return takes.ToArray();
        }

        internal Take[] GetLastTakes(Track track)
        {
            return GetLastTakes(track.Events);
        }

        internal Take[] GetLastTakes(TrackEvents events)
        {
            IEnumerable<Take> takes = events.Select(e => GetLastTake(e));
            return takes.ToArray();
        }

        internal Takes GetTakes(TrackEvent trackEvent)
        {
            return trackEvent.Takes;
        }

        internal Take GetFirstTake(TrackEvent trackEvent)
        {
            return trackEvent.Takes[0];
        }

        internal Take GetLastTake(TrackEvent trackEvent)
        {
            return trackEvent.Takes[trackEvent.Takes.Count - 1];
        }

        internal Take[] GetVideoTakes()
        {
            VideoTrack selected = SelectedVideoTrack();
            return GetFirstTakes(selected.Events);
        }

        internal Take[] GetAudioTakes()
        {
            AudioTrack selected = SelectedAudioTrack();
            return GetFirstTakes(selected.Events);
        }

        internal Media[] GetMediaList(VideoTrack track)
        {
            return GetMediaList(track.Events);
        }

        internal Media[] GetMediaList(AudioTrack track)
        {
            return GetMediaList(track.Events);
        }

        internal Media[] GetVideoMediaList()
        {
            VideoTrack selected = SelectedVideoTrack();
            return GetMediaList(selected.Events);
        }

        internal Media[] GetAudioMediaList()
        {
            AudioTrack selected = SelectedAudioTrack();
            return GetMediaList(selected.Events);
        }

        internal Media[] GetMediaList(TrackEvents events)
        {
            // テイクは考慮しない
            IEnumerable<Media> mediaList = events.Select(e => e.Takes[0].Media);
            return mediaList.ToArray();
        }

        internal OFXStringParameter GetOFXStringParameter(Media media, bool retNull = true)
        {
            foreach (OFXParameter param in media.Generator.OFXEffect.Parameters)
            {
                if (param.ParameterType == OFXParameterType.String)
                {
                    return (OFXStringParameter)param;
                }
            }
            if (retNull) { return null; }
            throw new VegasHelperNotFoundOFXParameterException();
        }

        internal OFXStringParameter[] GetOFXStringParameters(Media[] mediaList, bool retNull = true)
        {
            return mediaList.Select(m => GetOFXStringParameter(m, retNull)).ToList().ToArray();
        }

        internal OFXStringParameter[] GetOFXStringParameters(VideoTrack track, bool retNull = true)
        {
            Media[] mediaList = GetMediaList(track.Events);
            return GetOFXStringParameters(mediaList, retNull);
        }

        /// <summary>
        /// 選択したビデオトラックから、メディジェネレータ字幕のパラメータの配列を取得する
        /// ビデオトラックを選択していなければnullを返す
        /// </summary>
        /// <returns>選択したビデオトラックから得られたメディジェネレータ文字列パラメータの配列、もしくはnull</returns>
        internal OFXStringParameter[] GetOFXStringParameters(bool retNull = true)
        {
            VideoTrack selected = SelectedVideoTrack();
            return GetOFXStringParameters(selected, retNull);
        }

        public string GetOFXParameterString(OFXStringParameter param)
        {
            return param.GetValueAtTime(BaseTimecode);
        }

        public string GetOFXParameterString(Media media)
        {
            OFXStringParameter param = GetOFXStringParameter(media);
            return GetOFXParameterString(param);
        }

        internal string[] GetOFXParameterStrings(OFXStringParameter[] parameters)
        {
            return parameters.Select(p => GetOFXParameterString(p)).ToArray();
        }

        internal string[] GetOFXParameterStrings()
        {
            VideoTrack selected = SelectedVideoTrack();
            OFXStringParameter[] ofxParams = GetOFXStringParameters(selected);
            return GetOFXParameterStrings(ofxParams);
        }

        internal void SetStringIntoOFXParameter(OFXStringParameter param, string value)
        {
            param.SetValueAtTime(BaseTimecode, value);
        }

        /// <summary>
        /// 文字列の配列の内容を、メディジェネレータOFXの文字列パラメータ配列の各要素に設定する。
        /// 各引数の要素数が同じでないと処理しない。
        /// </summary>
        /// <param name="ofxParams">メディジェネレータOFXの文字列パラメータの配列</param>
        /// <param name="values">設定する文字列（RTF）の配列</param>
        internal void SetStringsIntoOFXParameters(OFXStringParameter[] ofxParams, string[] values)
        {
            if (ofxParams.Length != values.Length) { return; }
            for (int i = 0; i < ofxParams.Length; i++)
            {
                SetStringIntoOFXParameter(ofxParams[i], values[i]);
            }
        }

        internal OFXRGBAParameter GetTextRGBAParameter(Media media, bool retNull = true)
        {
            foreach (OFXParameter param in media.Generator.OFXEffect.Parameters)
            {
                if (param.ParameterType == OFXParameterType.RGBA &&
                    param.Name == "TextColor")
                {
                    return (OFXRGBAParameter)param;
                }
            }
            if (retNull) { return null; }
            throw new VegasHelperNotFoundOFXParameterException();
        }

        internal OFXRGBAParameter GetOutlineRGBAParameter(Media media, bool retNull = true)
        {
            foreach (OFXParameter param in media.Generator.OFXEffect.Parameters)
            {
                if (param.ParameterType == OFXParameterType.RGBA &&
                    param.Name == "OutlineColor")
                {
                    return (OFXRGBAParameter)param;
                }
            }
            if (retNull) { return null; }
            throw new VegasHelperNotFoundOFXParameterException();
        }

        internal OFXDoubleParameter GetOutlineWidthParameter(Media media, bool retNull = true)
        {
            foreach (OFXParameter param in media.Generator.OFXEffect.Parameters)
            {
                if (param.ParameterType == OFXParameterType.Double &&
                    param.Name == "OutlineWidth")
                {
                    return (OFXDoubleParameter)param;
                }
            }
            if (retNull) { return null; }
            throw new VegasHelperNotFoundOFXParameterException();
        }

        internal OFXRGBAParameter[] GetTextRGBAParameters(Media[] mediaList, bool retNull = true)
        {
            return mediaList.Select(m => GetTextRGBAParameter(m, retNull)).ToList().ToArray();
        }

        internal OFXRGBAParameter[] GetTextRGBAParameters(VideoTrack track, bool retNull = true)
        {
            Media[] mediaList = GetMediaList(track.Events);
            return GetTextRGBAParameters(mediaList, retNull);
        }

        internal OFXRGBAParameter[] GetTextRGBAParameters(bool retNull = true)
        {
            VideoTrack selected = SelectedVideoTrack();
            return GetTextRGBAParameters(selected, retNull);
        }

        internal void SetRGBAParameter(OFXRGBAParameter param, OFXColor color)
        {
            param.SetValueAtTime(BaseTimecode, color);
        }

        internal void SetDoubleParameter(OFXDoubleParameter param, double value)
        {
            param.SetValueAtTime(BaseTimecode, value);
        }

        internal VegasDuration GetEventTime(TrackEvent trackEvent)
        {
            VegasDuration duration = new VegasDuration();
            duration.StartTime = trackEvent.Start;
            duration.Length = trackEvent.Length;
            return duration;
        }

        internal void SetEventTime(TrackEvent trackEvent, VegasDuration duration, double margin = 0.0f, bool adjustTakes = true)
        {
            Timecode start = duration.StartTime - new Timecode(margin);
            Timecode length = duration.Length + new Timecode(margin * 2);
            trackEvent.AdjustStartLength(start, length, adjustTakes);
        }

        internal void AddTrackEventGroup(TrackEvent src, TrackEvent dst)
        {
            // Vegas.Project.TrackEventGroups.Addメソッドを先に呼ばいないと、
            // group.Addする際に例外が発生する
            TrackEventGroup group = new TrackEventGroup(Vegas.Project);
            Vegas.Project.TrackEventGroups.Add(group);
            group.Add(src);
            group.Add(dst);
        }

        internal void AssignAudioTrackDurationToVideoTrack(VideoTrack videoTrack, AudioTrack audioTrack, double margin = 0.0f, bool adjustTakes = true, bool group = true)
        {
            TrackEvents videoEvents = videoTrack.Events;
            TrackEvents audioEvents = audioTrack.Events;

            if (videoEvents.Count != audioEvents.Count) { return; }

            // TrackEventsのまま処理をするとリストの内容が勝手に入れ替わって不具合の原因になるため、
            // 別のListを作ってそこにTrackEventを挿入する
            List<TrackEvent> tmpVideoEvents = RefillTrackEvents(videoEvents);
            List<TrackEvent> tmpAudioEvents = RefillTrackEvents(audioEvents);

            for (int i = 0; i < videoEvents.Count; i++)
            {
                VegasDuration duration = GetEventTime(audioEvents[i]);
                SetEventTime(tmpVideoEvents[i], duration, margin, adjustTakes);

                if (group) { AddTrackEventGroup(tmpAudioEvents[i], tmpVideoEvents[i]); }
            }
        }

        internal void AssignAudioTrackDurationToVideoTrack(string trackName, double margin = 0, bool adjustTakes = true, bool group = true)
        {
            VideoTrack videoTrack = SearchVideoTrackByName(trackName);
            AudioTrack audioTrack = SearchAudioTrackByName(trackName);
            AssignAudioTrackDurationToVideoTrack(videoTrack, audioTrack, margin, adjustTakes, group);
        }

        internal void AssignAudioTrackDurationToVideoTrack(double margin = 0, bool adjustTakes = true, bool group = true)
        {
            VideoTrack videoTrack = SelectedVideoTrack();
            AudioTrack audioTrack = SelectedAudioTrack();
            AssignAudioTrackDurationToVideoTrack(videoTrack, audioTrack, margin, adjustTakes, group);
        }

        internal void DeleteJimakuPrefix()
        {
            VideoTrack track = SelectedVideoTrack();
            DeleteJimakuPrefix(track);
        }
        internal void DeleteJimakuPrefix(string title)
        {
            VideoTrack track = SearchVideoTrackByName(title);
            DeleteJimakuPrefix(track);
        }

        internal void DeleteJimakuPrefix(VideoTrack track)
        {
            DeleteJimakuPrefix(track.Events);
        }

        internal void DeleteJimakuPrefix(TrackEvents trackEvents)
        {
            foreach (TrackEvent trackEvent in trackEvents)
            {
                DeleteJimakuPrefix(trackEvent);
            }
        }

        internal TrackEvent GetFirstEvent(TrackEvents events)
        {
            return events[0];
        }

        internal TrackEvent GetLastEvent(TrackEvents events)
        {
            return events[events.Count - 1];
        }

        /// <summary>
        /// 選択したトラック内のイベントの開始位置と流さを求める。
        /// 最初のイベントの開始位置から最後のイベントの終点までをその長さとする。
        /// また、引数としてマージンもセット可能
        /// </summary>
        /// <param name="margin">設定するマージン。初期値は0.0</param>
        /// <returns></returns>
        internal VegasDuration GetDuretionFromAllEventsInTrack(double margin = 0.0f)
        {
            Track selected = SelectedTrack();
            return GetDuretionFromAllEventsInTrack(selected, margin);
        }

        internal VegasDuration GetDuretionFromAllEventsInTrack(Track track, double margin = 0.0f)
        {
            TrackEvents events = GetEvents(track);
            return GetDuretionFromAllEventsInTrack(events);
        }

        internal VegasDuration GetDuretionFromAllEventsInTrack(TrackEvents events, double margin = 0.0f)
        {
            TrackEvent firstEvent = GetFirstEvent(events);
            TrackEvent lastEvent = GetLastEvent(events);

            Timecode singleMaraginTimecode = new Timecode(margin);
            Timecode doubleMaraginTimecode = new Timecode(margin * 2);

            VegasDuration duration = new VegasDuration();
            duration.StartTime = firstEvent.Start - singleMaraginTimecode;
            duration.Length = lastEvent.Start + lastEvent.Length - firstEvent.Start + doubleMaraginTimecode;
            return duration;
        }

        internal long GetLengthFromAllEventsInTrack()
        {
            Track selected = SelectedTrack();

            return GetLengthFromAllEventsInTrack(selected);
        }

        internal long GetLengthFromAllEventsInTrack(Track track)
        {
            VegasDuration duration = GetDuretionFromAllEventsInTrack(track);

            return duration.Length.Nanos;
        }

        internal void ExpandFirstVideoEvent(double margin = 0.0)
        {
            TrackEvents videoEvents = GetVideoEvents();
            TrackEvents audioEvents = GetAudioEvents();
            VegasDuration duration = GetDuretionFromAllEventsInTrack(audioEvents, margin);
            SetEventTime(GetFirstEvent(videoEvents), duration);
        }

        private List<TrackEvent> RefillTrackEvents(TrackEvents trackEvents)
        {
            List<TrackEvent> events = new List<TrackEvent>();
            foreach (TrackEvent e in trackEvents) { events.Add(e); }
            return events;
        }
    }
}
