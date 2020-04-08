# FFMPEG_Scripts
Quick-use VBScript tools to re-encode videos using FFMPEG

## Use:
Drag and drop the video(s) onto the vbscript file

The ffmpeg.exe executable must either be in the same folder as the .vbs file, in a \bin subfolder, or in the [system PATH](td0g.ca/r/ffmpeg)

## ReEncode Scripts
The scripts apply audio or video filters to the video

For compatibility, the [recommendations on the ffmpeg website](https://trac.ffmpeg.org/wiki/Encode/H.264#Compatibility) have been followed


## 2Pass_Normalization
The script automates FFMPEG's 2-pass audio normalization filter 

The output audio level can be adjusted by the [audioLevel variable in the script](https://github.com/td0g/FFMPEG_Scripts/blob/master/2Pass_Normalize.vbs#L18)
