import sys
import os
import pymiere
from pymiere import wrappers
from pymiere import exe_utils

def seconds_to_timecode(seconds, fps):
    '''
    将秒转为时间码，有时可能会有1帧的误差
    '''
    s = int(seconds)
    f = int((seconds - s)*fps)
    m, s = divmod(s, 60)
    h, m = divmod(m, 60)
    return ("%02d;%02d;%02d;%02d" % (h, m, s, f))


if not exe_utils.is_premiere_running():
    raise ValueError("Premiere is not running")

project_opened, sequence_active = wrappers.check_active_sequence(crash=False)
if not (project_opened and sequence_active):
    raise ValueError("No project or sequence is opened")


project = pymiere.objects.app.project
sequence = project.activeSequence
active_sequence_qe = pymiere.objects.qe.project.getActiveSequence()
fps = 1/(float(sequence.timebase)/wrappers.TICKS_PER_SECONDS)
fps = round(fps, 2) #保留 2 位小数
size = str(sequence.frameSizeHorizontal) + ' X ' + str(sequence.frameSizeVertical)
print('Current project path: {}'.format(project.path))
print("Current sequence: {}，{}, {} fps".format(sequence.name, size, fps))


video_clips = []
tracks = sequence.videoTracks
for track in tracks:
    clips = track.clips
    for clip in clips:
        video_clips.append({
            'track': track.name or track.id, # Premiere 2017 doesn't have 'name' property on tracks and clips
            # 'name': clip.name
            'path': clip.projectItem.getMediaPath() if not clip.isMGT() else 'motion graphics template',
            'start': seconds_to_timecode(clip.start.seconds, fps),
            'end': seconds_to_timecode(clip.end.seconds, fps),
            'src_in': seconds_to_timecode(clip.inPoint.seconds, fps),
            'src_out': seconds_to_timecode(clip.outPoint.seconds, fps),
            'duration': seconds_to_timecode(clip.duration.seconds, fps),
            'preview_path': os.path.abspath('.')+'\\frame_'+str(clip.nodeId)+'.jpg'
        })
        # preview_frame_time = video_clips[-1]['start']
        # active_sequence_qe.exportFrameJPEG(preview_frame_time, video_clips[-1]['preview_path'])
print(video_clips)

#交互应该放到命令行里，包括 excel 的输出目录设置以及显示的文字信息等
#需要将时间码转为时:分:秒:帧


