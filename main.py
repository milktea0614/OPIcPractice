#-*- coding-utf-8 -*-

import os, sys, random, datetime, playsound

from PyQt5.QtWidgets import *
from PyQt5 import uic
from openpyxl import load_workbook

from gtts import gTTS
import pyaudio, wave


"""
.exe를 만들기 위한 file path
"""
current_path = os.path.dirname(os.path.realpath(__file__))
surveyUI_path = os.path.realpath(os.path.join(current_path, "OPIc_survey.ui"))
answerUI_path = os.path.realpath(os.path.join(current_path, "OPIc_answer.ui"))
excel_path = os.path.realpath(os.path.join(current_path, "OPIcQuestion_20211222.xlsx"))
icon_path = os.path.realpath(os.path.join(current_path, "winicon.jpg"))

result_dir_path = ""

"""
OPIc 질문지가 저장되어 있는 Excel의 기초 정보
Excel의 정보가 변경 되면 하기의 변수 정의 값들도 변할 가능성 있음
"""
SHEET_SUMMARY_SURVEY_END_INDEX = 74 # Summary Sheet의 설문조사 주제 부분의 끝 Index+1
SHEET_SUMMARY_END_INDEX = 94 # Summary Sheet의 끝 Index+1
SHEET_DATA_END_INDEX = 412 #Data Sheet의 끝 Index+1

"""
상단에 있는 excel_path를 가지고 wordbook 정의
"""
workbook = load_workbook(excel_path)

# survey_answer list
collect_answer = []
all_question_list = []

# question index
question_index = 0

# 2021-12-22, milktea0614@naver.com
def _make_result_directory():
    """
    연습 결과물이 저장되는 폴더 생성
    :return: None
    """
    global result_dir_path

    current_time = datetime.datetime.now().strftime("%Y%m%d%H%M")
    windows_user_name = os.path.expanduser('~')
    result_dir_path = os.path.realpath(os.path.join(windows_user_name,"Desktop", "OPIcPracticeResult_"+current_time))

    if os.path.exists(result_dir_path) == False:
        os.mkdir(result_dir_path)
        print("[INFO] result directory :"+result_dir_path)

# 2021-12-20, milktea0614@naver.com
def _get_questions_statement(p_id, p_worksheet_name):
    """
    id 값을 받아서 data sheet에 있는 질문 문장을 반환
    :param p_id: 질문의 id 값
    :param p_worksheet_name : 질문 문장이 있는 worksheet 이름
    :return: id값과 연결되어 있는 영문 질문 문장
    """

    statement = ""
    worksheet = workbook[p_worksheet_name]
    for i in range(2, SHEET_DATA_END_INDEX):
        if int(worksheet['A' + str(i)].value) == int(p_id):
            if p_worksheet_name=="Data":
                statement = worksheet['F' + str(i)].value
            else:
                statement = worksheet['D' + str(i)].value
            break
    return str(statement)

# 2021-12-17, milktea0614@naver.com
def _get_theme_range(s_list, f_list, r_list):
    """
    설문조사 결과인 s_list와 그 결과에 속해있는 주제마다 지니고 있는 빈출도를 나타내는 f_list
    :param s_list: 설문조사 결과에서 선택된 Theme을 나열한 list
    :param f_list: 설문조사 결과에서 선택된 Theme마다 주어진 빈출도를 나열한 list
    :param r_list: 설문조사 결과에서 선택된 Theme마다 가지고 있는 질문 id들의 범위(콤보, random 의 묶음)을 나열한 list
    :return: range=콤보/Random, 주제명, pop된 s_list, pop된 f_list, pop된 r_list
    """
    t_choice = random.choices(s_list, weights=f_list)
    t_index = s_list.index(t_choice[0])

    range = r_list.pop(t_index)
    theme = s_list.pop(t_index)

    del f_list[t_index]

    return range, theme, s_list, f_list, r_list

# 2021-12-17, milktea0614@naver.com
def _get_sub_questions_list(p_range, p_theme, p_questionNum):
    """
    범위에 작성되어 있는 내용을 토대로 영문 질문 리스트를 뽑아 반환
    :param p_range: 범위
    :param p_theme: 주제
    :param p_questionNum: 뽑아야 내야 할 질문 갯수
    :return: 영문 질문 리스트
    """
    result = []
    select = random.choices(p_range)
    t_question_id = []

    if '-' not in select[0]: # Random
        # sheet 바꿔서 리스트 가져오기
        worksheet = workbook["Data"]
        for i in range(3, SHEET_DATA_END_INDEX):
            if worksheet['C' + str(i)].value == p_theme:
                t_question_id.append(worksheet['A' + str(i)].value)

        # random 하게 N개 고르기
        for j in range(0,p_questionNum):
            t_num = random.choices(t_question_id)
            t_str = _get_questions_statement(t_num[0], "Data")
            result.append(t_str)
            t_question_id.remove(t_num[0])

    else:  # 콤보 문제일 때
        t_question_id_list = select[0].split("-")
        for i in range(0, p_questionNum):
            t_str = _get_questions_statement(t_question_id_list[i], "Data")
            result.append(t_str)

    return result

# 2021-12-22, milktea0614@naver.com
def get_question_list(p_survey_answer_list):
    """
    15개의 영문 질문 리스트 생성 및 all_question_list에 저장
    :param p_survey_answer_list: 설문조사에서 선택한 주제 리스트
    :return: None
    """
    worksheet = workbook["Summary"]

    # 설문 영역 list
    frequency_list = []
    question_range = []

    for i in range(0, len(p_survey_answer_list)):
        for j in range(3,SHEET_SUMMARY_SURVEY_END_INDEX):
            t_index= 'B'+str(j)
            if worksheet[t_index].value==p_survey_answer_list[i]:
                frequency_list.append(worksheet['C'+str(j)].value)
                temp = (worksheet['D'+str(j)].value).split(',')
                question_range.append(temp)

    # 돌발 영역 list
    outbreak_list=[]
    outbreak_frequency_list=[]
    outbreak_range_list=[]

    for i in range(74,SHEET_SUMMARY_END_INDEX):
        if worksheet['B'+str(i)].value != "":
            outbreak_list.append(worksheet['B'+str(i)].value)
            outbreak_frequency_list.append(worksheet['C'+str(i)].value)
            temp = str(worksheet['D'+str(i)].value).split(',')
            outbreak_range_list.append(temp)

    q_statment = []
    # 0차 선택 - 자기소개
    q_statment.append(_get_questions_statement(0, "Data"))

    # 1차 선택 - 설문 Theme (2개)
    range1, theme1, p_survey_answer_list, frequency_list, question_range = _get_theme_range(p_survey_answer_list, frequency_list, question_range)
    q_statment.extend(_get_sub_questions_list(range1, theme1, 2))

    # 2차 선택 - 설문 Theme (3개)
    range2, theme2, p_survey_answer_list, frequency_list, question_range = _get_theme_range(p_survey_answer_list, frequency_list, question_range)
    q_statment.extend(_get_sub_questions_list(range2, theme2, 3))

    # 3차 선택 - 설문 or 돌발 (3개)
    t_case = random.choice(["설문", "돌발"])
    if t_case == "설문":
        range3, theme3, p_survey_answer_list, frequency_list, question_range = _get_theme_range(p_survey_answer_list, frequency_list, question_range)
    else:
        range3, theme3, outbreak_list, outbreak_frequency_list, outbreak_range_list = _get_theme_range(outbreak_list, outbreak_frequency_list,outbreak_range_list)
    q_statment.extend(_get_sub_questions_list(range3, theme3, 3))

    # 4차 선택 - 돌발 3개
    range4, theme4, outbreak_list, outbreak_frequency_list, outbreak_range_list = _get_theme_range(outbreak_list,outbreak_frequency_list,outbreak_range_list)
    q_statment.extend(_get_sub_questions_list(range4, theme4, 3))

    # 5차 선택 - 돌발 1개, 롤플레이 2개
    range5_1, theme5_1, outbreak_list, outbreak_frequency_list, outbreak_range_list = _get_theme_range(outbreak_list,outbreak_frequency_list,outbreak_range_list)
    q_statment.extend(_get_sub_questions_list(range5_1, theme5_1, 1))

    role_list = list(range(1,46))
    range_5_2 = random.choices(role_list)
    role_list.pop(range_5_2[0])
    range_5_3 = random.choices(role_list)
    q_statment.append(_get_questions_statement(range_5_2[0],"Roleplay"))
    q_statment.append(_get_questions_statement(range_5_3[0],"Roleplay"))

    workbook.close()

    global all_question_list
    all_question_list = q_statment

# 2021-12-22, milktea0614@naver.com
def _make_question_audio_txt_files(p_list):
    """
    영문으로 작성된 질문 리스트를 tts를 이용해 오디오로 저장
    :param p_list 영문으로 된 질문 리스트
    :return: None
    """
    # for script files
    script_path = os.path.realpath(os.path.join(result_dir_path, "questions.txt"))
    script_f = open(script_path, 'w')
    print("[INFO] Create script file (" + script_path + ")")

    for i in range(0,len(p_list)):
        m_gtts = gTTS(text=p_list[i], lang='en')
        tts_file_name = os.path.join(current_path,"question"+str(i)+".mp3")
        m_gtts.save(tts_file_name)
        script_f.write(p_list[i]+"\n\n")
        print("[INFO] Create "+str(i)+" audio file ("+tts_file_name+")")

    script_f.close()

# 2021-12-22, milktea0614@naver.com
def _remove_question_audio_files():
    """
    오디오로 저장된 질문 파일 지우기
    :return: None
    """
    for i in range(0, len(all_question_list)):
        t_path = os.path.realpath(os.path.join(current_path,"question"+str(i)+".mp3"))
        if os.path.isfile(t_path):
            os.remove(t_path)


class SurveyWindow(QDialog):
    """
    설문조사를 담당하는 윈도우와 관련된 UI와 함수의 연결과 UI 출력 담당
    """
    def __init__(self):
        super(SurveyWindow, self).__init__()
        uic.loadUi(surveyUI_path, self)

        # Connect included Widgets
        self.BTN_startTest.clicked.connect(self.startTest)


    def startTest(self):
        """
        설문조사 화면에서 Start 버튼을 클릭했을때, 수행하는 동작
        :return: None
        """
        try:
            able, count = self.collect_survey_data()
            if able > 0:
                self.BTN_startTest.setText("Loading...")
                self.BTN_startTest.setEnabled(False)
                self.repaint()

                get_question_list(collect_answer)
                _make_result_directory()
                _make_question_audio_txt_files(all_question_list)

                self.close()

                answer_Window.show()

            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                if count < 12:
                    msg.setText("설문조사에서 체크한 영역이 부족합니다. ("+str(count)+"/12)")
                else:
                    msg.setText("설문조사에서 누락 된 영역이 있습니다.")

                msg.setWindowTitle("테스트 시작 불가")
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()
        except Exception as ex:
            print(ex)
            _remove_question_audio_files()
            exit()

    # 2021-12-16
    def collect_survey_data(self):
        """
        Survey GUI에서 선택된 Widget 정보를 가지고 와서 list로 저장
        :return: None
        """
        count = 0
        able_test = 0
        collect_answer.clear()

        # Part 1.
        if self.RB_employee.isChecked():
            collect_answer.append("직장인")
            able_test = 1
        elif self.RB_student.isChecked():
            collect_answer.append("학생")
            able_test = 1
        elif self.RB_jobseeker.isChecked():
            able_test = 1

        # Part 2
        for i in range(0, self.survey_grid2.count()):
            if self.survey_grid2.itemAt(i).widget().isChecked():
                collect_answer.append(self.survey_grid2.itemAt(i).widget().text())
                count += 1

        # Part 3
        for i in range(0, self.survey_grid3.count()):
            if self.survey_grid3.itemAt(i).widget().isChecked():
                collect_answer.append(self.survey_grid3.itemAt(i).widget().text())
                count += 1

        # Part 4
        for i in range(0, self.survey_grid4.count()):
            if self.survey_grid4.itemAt(i).widget().isChecked():
                collect_answer.append(self.survey_grid4.itemAt(i).widget().text())
                count += 1

        # Part 5
        for i in range(0, self.survey_horiz.count()):
            if self.survey_horiz.itemAt(i).widget().isChecked():
                collect_answer.append(self.survey_horiz.itemAt(i).widget().text())
                count += 1

        if count > 11 and able_test > 0:
            return 1, count
        else:
            return 0, count

class AnswerWindow(QDialog):
    """
    질문 재생과 답변 녹음을 담당하는 윈도우와 관련된 UI와 함수의 연결과 UI 출력 담당
    """
    def __init__(self, chunk=1024, p_format=pyaudio.paInt16, channels=1, rate=44100, py=pyaudio.PyAudio()):
        super(AnswerWindow, self).__init__()
        uic.loadUi(answerUI_path, self)

        # Connect included Widgets
        self.BTN_start.clicked.connect(self.playQuestionAudio)
        self.BTN_complete.clicked.connect(self.completeQuestion)
        self.BTN_Next.clicked.connect(self.goToNextQuestion)

        # Record 관련
        self.CHUNK = chunk
        self.FORMAT = p_format
        self.CHANNELS = channels  # 녹음 장비 번호
        self.RATE = rate
        self.stop = 0
        self.frames = []
        self.p = py
        self.stream = self.p.open(format=self.FORMAT, channels=self.CHANNELS, rate=self.RATE, input=True,frames_per_buffer=self.CHUNK)

    def record_answer(self):
        """
        사용자가 녹음장치(1번장치)로 녹음하는 음성을 저장하는 함수
        :return: None
        """
        self.stop = 1
        self.frames = []
        stream = self.p.open(format=self.FORMAT, channels=self.CHANNELS, rate=self.RATE, input=True,frames_per_buffer=self.CHUNK)
        self.LB_status.setText("녹음")
        self.repaint()

        while self.stop==1:
            data = self.stream.read(self.CHUNK)
            self.frames.append(data)
            QApplication.processEvents() # UI freeze 방지

        stream.close()
        t_path = os.path.realpath(os.path.join(result_dir_path,"answer"+str(question_index)+".wav"))
        t_wave_file = wave.open(t_path, "wb")
        t_wave_file.setnchannels(self.CHANNELS)
        t_wave_file.setsampwidth(self.p.get_sample_size(self.FORMAT))
        t_wave_file.setframerate(self.RATE)
        t_wave_file.writeframes(b''.join(self.frames))
        t_wave_file.close()

        self.BTN_complete.setEnabled(False)
        self.BTN_Next.setEnabled(True)
        self.LB_status.setText("대기중")

        if question_index == 14:
            self.BTN_Next.setText("종료")

        self.repaint()


    def playQuestionAudio(self):
        """
        문제 재생 버튼 클릭 시, 작동하는 함수
        :return: None
        """
        self.LB_status.setText("문제 재생")
        self.BTN_complete.setEnabled(True)
        self.BTN_start.setEnabled(False)
        self.repaint()

        question_file = os.path.join(current_path, "question" + str(question_index) + ".mp3")
        playsound.playsound(question_file)

        self.record_answer()

    def completeQuestion(self):
        """
        종료 버튼 클릭 시, 작동하는 함수
        :return: None
        """
        self.stop = 0
        self.repaint()

    def goToNextQuestion(self):
        """
        Next 버튼 클릭시 작동하는 함수
        :return: None
        """
        global question_index

        #마지막 문제가 아닐 때
        if question_index < 14:
            question_index+=1
            self.LB_QuestionNum.setText(str(question_index+1))
            self.BTN_start.setEnabled(True)
            self.BTN_complete.setEnabled(False)
            self.BTN_Next.setEnabled(False)
        else:
            _remove_question_audio_files()
            os.startfile(result_dir_path)
            exit()

if __name__ == '__main__':
    apps = QApplication(sys.argv)

    survey_Window = SurveyWindow()
    answer_Window = AnswerWindow()

    # Widget
    survey_Window.show()

    # 프로그램 작동 코드
    apps.exec_()

