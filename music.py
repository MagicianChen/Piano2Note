from piano_transcription_inference import PianoTranscription, sample_rate, load_audio
import mido
import os
from pydub import AudioSegment
from pydub.silence import split_on_silence
import xlrd
import xlwt
import xlutils.copy
import copy


class pianoTrans():
    def __init__(self, wav_path):
        self.wav_path = wav_path

        self.wav_files = []

        for maindir, subdir, file_name_list in os.walk(self.wav_path):

            for filename in file_name_list:
                ext = os.path.splitext(filename)[1]

                if ext == '.wav':
                    self.wav_files.append(filename)
        self.wav_files = sorted(self.wav_files)

        self.match_dict = {48: 1, 49: 1,50: 2, 51: 2, 52: 3, 53: 4, 54:4, 55: 5, 56:5, 57:6, 58:6, 59: 7}

        self.transcriptor = PianoTranscription(device='cuda',
                                               checkpoint_path='./CRNN_note_F1=0.9677_pedal_F1=0.9186.pth')  # device: 'cuda' | 'cpu'


    def wav2midi(self, split_file):
        '''
        transfer the wav audio to midi note
        :param split_file: the processed file name
        :return: Each note identified in the audio corresponds to a number.
                Every 12 counts an octave, and 60 is the central C.
        '''
        (audio, _) = load_audio(split_file, sr=sample_rate, mono=True)

        transcribed_dict = self.transcriptor.transcribe(audio, split_file.split('.')[0] + '.mid')

        return sorted(transcribed_dict['est_note_events'], key = lambda x: float(x['onset_time']))

    def midi2note(self, midi_name):
        mid = mido.MidiFile(os.path.join(self.wav_path, midi_name))
        for i, track in enumerate(mid.tracks):
            for msg in track:
                print(dict(msg))
                exit(0)

    def split_wav(self, wav_path, save_dir):
        '''
        split the experimental audio according to the slient paragraph
        :param wav_path: the wave path
        :param save_dir: the output path for the splitted audio
        :return:
        '''
        sound = AudioSegment.from_mp3(wav_path)

        chunks = split_on_silence(sound,
                                  # must be silent for at least half a second,沉默半秒
                                  min_silence_len=4000,

                                  # consider it silent if quieter than -16 dBFS
                                  silence_thresh=-45,
                                  keep_silence=100,
                                  seek_step=2000,

                                  )
        for i in list(range(len(chunks)))[::-1]:
            if len(chunks[i]) <= 2000:
                chunks.pop(i)

        results = []
        for i, chunk in enumerate(chunks):
            chunk.export(os.path.join(save_dir, "{0}.wav".format(i)), format="wav")
            results.append(os.path.join(save_dir, "{0}.wav".format(i)))

        return results


    def create_xls(self, save_path, file_name):
        '''
        create an excel to write the results
        :param save_path: the save path
        :param file_name: the subject name
        :return:
        '''
        value = [["标准音符", "实验结果", "准确率"]]
        index = len(value)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("results")
        for i in range(0, index):
            for j in range(0, len(value[i])):
                sheet.write(i, j, value[i][j])
        workbook.save(os.path.join(save_path, file_name))
        print("{} 表格创建成功！".format(file_name.split('.')[0]))

    def write_xls(self, save_path, file_name, results):
        '''
        keep to write the results to the excel
        :param save_path:
        :param file_name:
        :param results:
        :return:
        '''
        index = len(results)  
        workbook = xlrd.open_workbook(os.path.join(save_path, file_name))
        sheets = workbook.sheet_names()
        worksheet = workbook.sheet_by_name(sheets[0])
        rows_old = worksheet.nrows
        new_workbook = xlutils.copy.copy(workbook)
        new_worksheet = new_workbook.get_sheet(0)
        for i in range(0, index):
            for j in range(0, len(results[i])):
                new_worksheet.write(i + rows_old, j, results[i][j])
        new_workbook.save(os.path.join(save_path, file_name))
        print("{} 实验结果写入完成！".format(file_name.split('.')[0]))


    def accuracy(self, ref, predict):
        '''
        Whether the predict note can match the answer exactly after controlling for error
        :param ref: the correct piano note
        :param predict: the output note from the subjects
        :return: whether the predict is correct or not
        '''

        flag = 1

        for note in ref:
            if note not in predict:
                flag = 0
                return str(flag)
            else:
                predict.remove(note)

        return str(flag)


    def normalize(self, output):
        for idx in range(len(output)):
            while output[idx]['midi_note'] < 48 or output[idx]['midi_note'] >= 60:
                if output[idx]['midi_note'] < 48:
                    output[idx]['midi_note'] += 12

                if output[idx]['midi_note'] >= 60:
                    output[idx]['midi_note'] -= 12

        return output


    def run(self):
        if not os.path.exists('output/output_xls'):
            os.mkdir('output/output_xls')
        for wav_file in self.wav_files:
            if os.path.exists(os.path.join('output/output_xls',  wav_file.split('.')[0] + '.xls')):
                continue

            if not os.path.exists(os.path.join('output', wav_file.split('.')[0])):
                os.mkdir(os.path.join('output', wav_file.split('.')[0]))

            split_files = self.split_wav(os.path.join(self.wav_path, wav_file), os.path.join('output', wav_file.split('.')[0]))
            # while len()
            split_files.sort(key = lambda x: int(x.split('.')[0].split('/')[-1]))
            self.create_xls('output/output_xls',  wav_file.split('.')[0] + '.xls')


            results = []
            for split_file in split_files:
                output = self.wav2midi(split_file)
                output = self.normalize(output)
                results.append([str(self.match_dict[output[i]['midi_note']]) for i in range(len(output))])

            processed_results = []
            for idx in range(len(results)):
                processed_results.append([''.join(results[idx][:7]), ''.join(results[idx][7:]), self.accuracy(results[idx][:7], results[idx][7:])])
                processed_results.append([''] * len(processed_results[-1]))  # add one empty line
            self.write_xls('output/output_xls',  wav_file.split('.')[0] + '.xls', processed_results)

            # self.write_xls()


if __name__ == "__main__":
    file_path = './wav_audio'  # the audio path

    transcriptor = pianoTrans(file_path)

    transcriptor.run()