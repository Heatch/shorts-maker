import pandas as pd
import azure.cognitiveservices.speech as speechsdk
from moviepy.editor import VideoFileClip, CompositeAudioClip, AudioFileClip, vfx, afx, CompositeVideoClip, TextClip
from moviepy.video.tools.subtitles import SubtitlesClip
import os
import whisper_timestamped as whisper
import json

# constants
fps = 24
sheet = "spreadsheet.xlsx"
voiceName = 'en-CA-LiamNeural'
jsonFile = "transcription.json"
apiKey = "YOUR_API_KEY"

speech_config = speechsdk.SpeechConfig(subscription=apiKey, region="eastus")
speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Riff48Khz16BitMonoPcm)
speech_config.speech_synthesis_voice_name=voiceName
speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=None)
model = whisper.load_model("tiny", device="cpu")

def buildShort(text, title):

    speech_synthesis_result = speech_synthesizer.speak_text_async(text).get()
    stream = speechsdk.AudioDataStream(speech_synthesis_result)
    stream.save_to_wav_file("audio.wav")
    audio = whisper.load_audio("audio.wav")
    result = whisper.transcribe(model, audio, language="en")

    with open(jsonFile, 'w') as fp:
        json.dump(result, fp, indent=4)

    with open(jsonFile, 'r') as f:
        data = json.load(f)

    all_words = []
    for segment in data['segments']:
        for word in segment['words']:
            all_words.append(word)

    main_audio = AudioFileClip("audio.wav")
    duration = main_audio.duration
    music = AudioFileClip("music.mp3").subclip(0, duration)
    music = music.volumex(0.3).fx(afx.audio_fadeout, 0.5)
    final_audio = CompositeAudioClip([main_audio, music])
    mainClip = VideoFileClip("backdrop.mp4").subclip(0, duration)

    generator = lambda txt: TextClip(txt, font='Arial', fontsize=150, color='white', stroke_color='black', stroke_width=3)
    subs = []
    for word in all_words:
        subs.append(((word['start'], word['end']), word['text']))

    subtitles = SubtitlesClip(subs, generator)
    result = CompositeVideoClip([mainClip, subtitles.set_pos(('center', 'center'))]).fx(vfx.fadeout, 0.5)
    result = result.set_audio(final_audio)

    result.write_videofile(title + ".mp4", fps=fps, codec="libx264", audio_codec="aac", threads=8, preset="ultrafast")
    os.remove("audio.wav")

df = pd.read_excel(sheet, engine='openpyxl', dtype=object, header=None)
cellTexts = df.values.tolist()
cellTexts = cellTexts[1:]

for cellText in cellTexts:
    buildShort(cellText[0], cellText[1])
