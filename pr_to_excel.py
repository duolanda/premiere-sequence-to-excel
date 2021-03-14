import sys
import os
import time
import argparse
import pymiere
from pymiere import wrappers
from pymiere import exe_utils
import xlsxwriter


def seconds_to_timecode(seconds, fps):
    '''
    将秒转为时间码，有时可能会有1帧的误差
    '''
    s = int(seconds)
    f = int((seconds - s)*fps)
    m, s = divmod(s, 60)
    h, m = divmod(m, 60)
    return ("%02d;%02d;%02d;%02d" % (h, m, s, f))

def scale_factor(width, height):
    '''
    根据视频大小确定表格中图片的缩放比例
    '''
    x = (1920 / width) * 0.167
    y = (1080 / height) * 0.02
    x, y = round(x, 3), round(y, 3)
    return x, y



def get_clips_info(is_middle):
    '''
    将视频片段的相关信息都以字典形式存到 video_clips 数组中
    '''
    global size_x, size_y
    
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
                'middle': seconds_to_timecode((clip.start.seconds+clip.end.seconds)/2, fps),
                'src_in': seconds_to_timecode(clip.inPoint.seconds, fps),
                'src_out': seconds_to_timecode(clip.outPoint.seconds, fps),
                'duration': seconds_to_timecode(clip.duration.seconds, fps),
                'preview_path': os.path.abspath('.')+'\\frame_'+str(clip.nodeId)+'.jpg',
            })
            if is_middle:
                preview_frame_time = video_clips[-1]['middle']
            else:
                preview_frame_time = video_clips[-1]['start']
            active_sequence_qe.exportFrameJPEG(preview_frame_time, video_clips[-1]['preview_path'])

    print('Loading preview images')
    while(True):
        img_list = []
        for clip in video_clips: 
            img_list.append(os.path.isfile(clip['preview_path']))
        if not (False in img_list):
            print('Images loading completed')
            time.sleep(1)
            break

    #如果有 PIL 的话用它减小一下图片尺寸
    if size_x >= 1920:
        try:
            from PIL import Image
            size_x = 960
            size_y = 540 
            for clip in video_clips: 
                img = Image.open(clip['preview_path'])
                out = img.resize((960, 540),Image.ANTIALIAS) 
                out.save(clip['preview_path'])
        except:
            pass

    return video_clips


def write_into_excel(video_clips, output_path):
    '''
    将相关信息写入 excel 文件
    '''
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 45) #图片列宽一些
    worksheet.set_column('D:H', 12)

    # 首行
    worksheet.write('A1', 'Preview')
    worksheet.write('B1', 'Track')
    worksheet.write('C1', 'Path')
    worksheet.write('D1', 'Start')
    worksheet.write('E1', 'End')
    worksheet.write('F1', 'Src_in')
    worksheet.write('G1', 'Src_out')
    worksheet.write('H1', 'Duration')

    # 图片缩放系数
    x, y = scale_factor(size_x, size_y)

    for i in range(len(video_clips)): 
        clip = video_clips[i]
        worksheet.insert_image(i+1, 0, clip['preview_path'], {'x_scale': x, 'y_scale': y, 'positioning': 1}) #positioning 1 图片位置和大小都随表格变化
        worksheet.write(i+1, 1, clip['track'])
        worksheet.write(i+1, 2, clip['path'])
        worksheet.write(i+1, 3, clip['start'])
        worksheet.write(i+1, 4, clip['end'])
        worksheet.write(i+1, 5, clip['src_in'])
        worksheet.write(i+1, 6, clip['src_out'])
        worksheet.write(i+1, 7, clip['duration'])

    workbook.close()

    for clip in video_clips: 
        os.remove(clip['preview_path'])


    print('Excel export succeed')


if __name__ == "__main__":
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
    size_x = sequence.frameSizeHorizontal
    size_y = sequence.frameSizeVertical
    print('Current project path: {}'.format(project.path))
    print("Current sequence: {}，{}, {} fps".format(sequence.name, size, fps))





    parser = argparse.ArgumentParser(
            description='Python script for exporting Adobe Premiere sequence clips into excel file.')
            
    parser.add_argument('-o','--out',
                        default = os.path.abspath('.')+'\\sequence_clip.xlsx', 
                        help = "output excel folder")
    parser.add_argument('-m','--middle',
                        type = int,
                        default = 1, 
                        help = "0 to capture preview from the first frame of clip, "
                               "1 to capture preview from the middle frame of clip.")

    args = parser.parse_args()


    video_clips = get_clips_info(args.middle)
    write_into_excel(video_clips, args.out)


    
