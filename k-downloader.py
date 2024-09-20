# -*- coding: utf-8 -*-
# SPDX-License-Identifier: MIT

import re
from subprocess import Popen, PIPE, check_output
import ctypes
import time
import wx
import wx.dataview as dv
import wx.html
import wx.lib.wxpTag
import threading
from threading import Thread
import os
import sys
import pickle
import win32con
import win32gui
import win32api
import win32process
import win32com.client
import requests
import clipboard
import math
import psutil


TITLE = 'K-Downloader'

def is_running(name):
    for task in check_output(['tasklist'], creationflags=0x08000000).decode('utf-8', 'ignore').split("\r\n"):
        m = re.match("(.+?) +(\d+) (.+?) +(\d+) +(\d+.* K).*", task)
        if m is not None:
            if m.group(1) == name:
                # print(m.group(1))  # 파일명 (K-Downloader.exe)
                # print(m.group(2))  # PID
                # print(m.group(3))  # 유형 (Console, Services)
                # print(m.group(5))  # 사용중인 메모리 양
                return True

    return False


URL = 'https://www.youtube.com/watch?v=yZeNfaIK7Nw'  # 경복궁
VERSION = '2024.09.09'
PYTHON = '3.10.10'
WXPYTHON = '4.2.0'
FFMPEG2 = '2024-06-21-git-d45e20c37b'
PYINSTALLER = '5.9.0'

FFMPEG = '.\\ffmpeg.exe'
YT_DLP = 'yt-dlp.exe'
YT_DLP_PATH = f'.\\{YT_DLP}'


class Help(wx.Dialog):
    def __init__(self, parent, arg):
        title = '도움말' if arg != 9 else f'{TITLE} 정보'
        wx.Dialog.__init__(self, parent, -1, title=title)
        html = wx.html.HtmlWindow(self, -1, size=(440, -1))
        text = ''
        if arg == 1:
            text += """<html><body>
<h3>포맷 선택</h3>
<table width="440" cellspacing="0" cellpadding="0" border="0">
<tr>
    <td width="170">포맷</td><td width="100">파일크기</td><td></td>
</tr>
</table>
<hr>
<table width="440" cellspacing="0" cellpadding="0" border="0">
<tr>
    <td width="170">1920x1012 (mp4, 3368k)</td>
    <td width="100">55.25MiB</td>
    <td>가로 1920px, 세로 1012px
    <br>파일 확장자 .mp4
    <br><strong>비디오비트전송률</strong>(=영상 비트레이트) 3368k
    <br>파일 크기 55.25MiB</td>
</tr>
<tr>
    <td colspan="3">- 선택 시 '오디오 추가' 목록이 나타나면 영상만 있고 <strong>음성은 없는</strong>(Video only, no audio) 것임.
    <br>⇒ 오디오를 추가하면 <strong>음성 있는</strong> 동영상이 생성됨.</td>
</tr>
</table>
<hr>
<table width="440" cellspacing="0" cellpadding="0" border="0">
<tr>
    <td width="170">640x348 (mp4, <strong>591k</strong>)</td>
    <td width="100">9.69MiB</td>
    <td>가로 640px, 세로 348px
    <br>파일 확장자 .mp4
    <br><strong>총비트전송률</strong>(=영상+음성 비트레이트) 591k
    <br>파일 크기 9.69MiB</td>
</tr>
<tr>
    <td colspan="3">- 선택 시 '오디오 추가' 목록이 안 나타나면 영상과 음성 <strong>둘 다 있는</strong>(Both video and audio) 것임.</td>
</tr>
</table>
<hr>
<table width="440" cellspacing="0" cellpadding="0" border="0">
<tr>
    <td width="170">1920x1080 (<strong>m3u8</strong>, 4561k)</td>
    <td width="100"></td>
    <td>가로 1920px, 세로 1080px
    <br>파일 확장자 <strong>.m3u8</strong>
    <br>총비트전송률(=영상+음성 비트레이트) 4561k
    <br>파일 크기 알 수 없음</td>
</tr>
<tr>
    <td colspan="3">- 확장자가 .m3u8이면 통신 프로토콜이 <strong>HLS</strong>(HTTP 라이브 스트리밍)인 것임.</td>
</tr>
</table>
<hr>
<a href="https://ko.wikipedia.org/wiki/mp4">mp4</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/m4a">m4a</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/WebM">webm</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/HTTP_라이브_스트리밍">HLS</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/GiB">GiB</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/MiB">MiB</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/비트레이트">비트레이트</a><br><br>
<a href="https://ko.wikipedia.org/wiki/8K_해상도">8K Ultra HD</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/4K_해상도">4K Ultra HD</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/1080p">Full HD</a>"""

        elif arg == 2:
            text += """<html><body>
<h3>오디오 추가</h3>
<p>영상만 있고 <strong>음성은 없는</strong>(Video only, no audio) 포맷에 오디오 스트림을 추가합니다. 
<p>이때, 영상파일 확장자와 음성파일 확장자가 어떤가에 따라 다운로드 파일의 확장자도 달라집니다.
<p>
<table width="440" cellspacing="0" cellpadding="0" border="0">
<tr>
    <td colspan="3"><hr></td>
</tr>
<tr>
    <td width="120">영상(프로토콜)</td><td width="120">&nbsp;&nbsp;음성(프로토콜)</td><td>&nbsp;&nbsp;&nbsp;&nbsp;다운로드 파일</td>
</tr>
<tr>
    <td colspan="3"><hr></td>
</tr>
<tr>
    <td>mp4(https)<br></td><td>+ m4a(https)<br></td><td>=> mp4<br></td>
</tr>
<tr>
    <td>mp4(https)<br></td><td>+ webm(https)<br></td><td>=> mkv<br></td>
</tr>
<tr>
    <td>mp4(https)<br></td><td>+ mp4(m3u8)<br></td><td>=> mp4<br></td>
</tr>
<tr>
    <td>webm(https)<br></td><td>+ m4a(https)<br></td><td>=> mkv<br></td>
</tr>
<tr>
    <td>webm(https)<br></td><td>+ webm(https)<br></td><td>=> webm<br></td>
</tr>
<tr>
    <td>mp4(m3u8)<br></td><td>+ webm(https)<br></td><td>=> webm<br></td>
</tr>
<tr>
    <td>mp4(m3u8)<br></td><td>+ m4a(https)<br></td><td>=> mp4<br></td>
</tr>
<tr>
    <td>mp4(m3u8)</td><td>+ mp4(m3u8)</td><td>=> mp4</td>
</tr>
</table>
<hr>
<a href="https://ko.wikipedia.org/wiki/mp4">mp4</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/m4a">m4a</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/WebM">webm</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/.mkv">mkv</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/컴프레서_(오디오)">DRC</a>"""

        elif arg == 3:
            text += """<html><body>
<h3>후처리:리먹싱(=>MP4)</h3>
<h4>1. 리먹싱(Remuxing)이란?</h4>
<p>- "리멀티플렉싱(remultiplexing)"의 줄임말인 리먹싱은 하나의 멀티미디어 파일(예: MP4, MKV 또는 AVI)의 내용을 가져와 다른 파일 형식으로 다시 래핑하는 비디오 처리 방법입니다. 이 프로세스는 어떤 식으로든 인코딩 원설정을 변경하거나 오디오, 비디오 또는 자막 스트림을 수정하지 않습니다. 대신 이러한 요소를 한 컨테이너 형식에서 다른 컨테이너 형식으로 전송하기만 하면 됩니다.
<br>
<h4>2. 리먹싱의 이점</h4>
<p>- <strong>호환성</strong>: 리먹싱을 통해 사용자는 파일을 특정 장치 또는 미디어 플레이어와 호환되는 형식으로 변환할 수 있음. 예를 들어 MKV 파일을 MP4로 변환하면 iOS 기기에서 재생할 수 있음.
<p>- <strong>파일 크기 최적화</strong>: 보다 효율적인 컨테이너 형식을 선택하면 리먹싱을 통해 비디오 품질을 손상시키지 않으면서 파일 크기가 작아질 수 있음. 이 기능은 저장 공간을 절약하거나 파일을 스트리밍에 더 적합하게 만드는 데 특히 유용함.
<p>- <strong>품질 보존</strong>: 리먹싱에는 재인코딩이 포함되지 않으므로 원본 비디오 및 오디오 품질을 보존함. 이는 미세한 품질 손실도 눈에 띌 수 있는 고품질 비디오 콘텐츠에 특히 중요함.
<p>- <strong>자막 지원</strong>: 리먹싱을 사용하여 자막을 한 형식에서 다른 형식으로 전송하여 변환 프로세스 중에 자막이 그대로 유지되도록 할 수 있음.
<p>- <strong>재압축 없음</strong>: 비디오 및 오디오 스트림을 디코딩하고 다시 인코딩하는 트랜스코딩과 달리 리먹싱은 재압축을 완전히 방지함. 이렇게 하면 시간이 절약될 뿐만 아니라 아티팩트가 발생하거나 비디오 품질이 저하될 위험도 줄어듦.
<p>- <strong>다양성</strong>: 리먹싱은 다양한 파일 형식에 적용할 수 있어 비디오 처리 작업을 위한 다목적 도구임. Blu-ray 디스크를 디지털 형식으로 변환해야 하거나 단순히 비디오 파일의 컨테이너 형식을 변경해야 하는 경우 리먹싱을 사용하면 됨.
<br>
<p>인용: <a href="https://mps.live/blog/details/what-is-remuxing">https://mps.live/blog/details/what-is-remuxing</a>"""

        elif arg == 9:
            text += f"""<html><body>
<h3>{TITLE}</h3><strong>버전</strong> {VERSION}<br><br>
<p><strong><a href="https://github.com/yt-dlp/yt-dlp">yt-dlp</a></strong> {parent.ytdlp_current_version}
<p><strong><a href="https://www.ffmpeg.org/">FFmpeg</a></strong> {FFMPEG2}
<p><strong><a href="https://wxpython.org/">wxPython</a></strong> {WXPYTHON}
<p><strong><a href="https://www.python.org/">Python</a></strong> {PYTHON}
<p><strong><a href="https://pyinstaller.org/">PyInstaller</a></strong> {PYINSTALLER}<br><br>
<p>HS Kang
<hr>
<a href="https://namu.wiki/w/yt-dlp">yt-dlp</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/FFmpeg">FFmpeg</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/WxPython">wxPython</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://ko.wikipedia.org/wiki/Python">Python</a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="https://en.wikipedia.org/wiki/Pyinstaller">PyInstaller</a>"""

        text += """<p style="text-align:right"><wxp module="wx" class="Button">
    <param name="label" value="닫기">
    <param name="id"    value="ID_OK">
</wxp></p>
</body></html>"""

        html.SetPage(text)
        ir = html.GetInternalRepresentation()
        html.SetSize((ir.GetWidth() + 25, ir.GetHeight() + 0))
        self.SetClientSize(html.GetSize())
        self.CentreOnParent(wx.BOTH)
        html.Bind(wx.html.EVT_HTML_LINK_CLICKED, self.onevtlinkclicked)

    @staticmethod
    def onevtlinkclicked(link):
        url = link.GetLinkInfo().GetHref()
        os.startfile(url)

    def __del__(self):
        pass
        # frame.SetFocus()


def evt_result(win, func):
    win.Connect(-1, -1, -1, func)


def getseconds(s):
    h, m, s = s.split(':')
    sec = int(h) * 3600 + int(m) * 60 + float(s)
    return sec


def get_version_number(file_path):
    file_information = win32api.GetFileVersionInfo(file_path, "\\")
    ms_file_version = file_information['FileVersionMS']
    ls_file_version = file_information['FileVersionLS']

    return f'{win32api.HIWORD(ms_file_version)}.' \
           f'{win32api.LOWORD(ms_file_version):02d}.' \
           f'{win32api.HIWORD(ls_file_version):02d}'


class ResultEvent(wx.PyEvent):
    def __init__(self, data):
        wx.PyEvent.__init__(self)
        self.SetEventType(-1)
        self.data = data


class WorkerThread(Thread):
    def __init__(self, parent):
        Thread.__init__(self)
        self.parent = parent
        self.table_started = False
        self.pipe_pos = []
        self.del_origin_cnt = 0
        self.table_lines = []
        
    def run(self):
        parent = self.parent
        # parent.cancelled = False
        parent.task_done = False
        parent.progress_count = 0
        parent.gauge.SetValue(parent.progress_count)
        s = '[정보 추출] 준비 중입니다. 잠깐만 기다려주세요..'  # if parent.task == 'extract':
        if parent.task == 'download':
            parent.percent_last = -1
            s = '[동영상 다운로드] 준비 중입니다. 잠깐만 기다려주세요..'

        parent.status.SetLabel(s)

        if parent.task == 'extract':
            self.extract_it()

        elif parent.task == 'download':
            self.download_it()

    def abort(self, evt=None):
        suffix = '' if evt else '2'
        wx.PostEvent(self.parent, ResultEvent(f'{self.parent.task}-cancelled{suffix}'))
        self.raise_exception()

    def raise_exception(self):
        thread_id = self.get_id()
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id,
                                                         ctypes.py_object(SystemExit))
        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0)
            print('예외 발생 실패')

    def get_id(self):
        if hasattr(self, '_thread_id'):
            return self._thread_id

        for t in threading.enumerate():
            if t is self:
                return t.native_id

    def extract_it(self):
        parent = self.parent
        url = parent.txtURL.GetValue()
        parent.is_live = False
        parent.formats_data = {}
        self.table_lines = []

        # 유튜브, 네이버TV, 비메오
        if 'youtu.be/' in url or 'youtube.com/watch?v=' in url:
            parent.host = 'youtube'
            url = re.sub(r'[&?]list=.*', '', url)
        elif 'tv.naver.com/v/' in url:
            parent.host = 'naver'
        elif 'vimeo.com/' in url:
            parent.host = 'vimeo'
        elif 'afreecatv.com/' in url:
            parent.host = 'afreecatv'
        elif 'facebook.com/' in url:
            parent.host = 'facebook'
        elif 'tv.kakao.com/' in url:
            parent.host = 'kakao'
        else:
            parent.host = 'etc'

        cmd = f'powershell & "{YT_DLP_PATH}" --list-formats -j --no-quiet "{url}"'.split() + \
              '2>&1 | % ToString | Tee-Object out.txt'.split()

        # print('task =>', parent.task, ', cmd =>', ' '.join(cmd))

        parent.proc = Popen(cmd, stdin=PIPE, stdout=PIPE, stderr=PIPE, creationflags=0x08000000)
        while parent.proc.poll() is None:
            self.checkproc_extract()

        if parent.task_done:
            wx.PostEvent(parent, ResultEvent('extract-finished'))

    def checkproc_extract(self):
        parent = self.parent
        s = str(parent.proc.stdout.readline())
        #print(s)
        if s == "b''":
            return
        s = s.replace("b'", "").replace("\\r\\n'", "")
        s = s.replace('b"', '').replace('\\r\\n"', '')
        print(s)

        if 'getaddrinfo failed' in s or \
            'Unable to download webpage' in s or \
            'Unable to download API page' in s or \
            'Downloading iframe API JS' in s:

            msg = '인터넷 연결상태를 확인해주세요.'

            s = re.sub('.*Unable to download webpage.*', '웹페이지를 다운로드할 수 없습니다.', s)
            s = re.sub('.*getaddrinfo failed.*', 'DNS 조회에 실패했습니다.', s)
            s = re.sub('.*Unable to download API page.*', 'API 페이지를 다운로드할 수 없습니다.', s)
            s = re.sub('.*Downloading iframe API JS.*', 'iframe API JS 다운로드', s)

            print(s)
            parent.status.SetLabel(s)
            parent.gauge.SetValue(0)
            wx.MessageBox(f'{msg}\n\n{s}', TITLE, style=wx.ICON_ERROR)
            self.abort()
            return

        if s == "b''":
            return

        if re.search('n =.*player =', s):
            return

        s = s.replace("b'", "").replace("\\r\\n'", "")
        # print(s)

        if s.startswith('ERROR:'):
            print(s)
            msg = '에러:'

            s = re.sub('please report.*', '', s)
            s = re.sub('This live event will begin in a few moments', '이 라이브 이벤트는 잠시 후에 시작됩니다', s)
            s = re.sub('This live event will begin in (\d+) minutes', r'이 라이브 이벤트는 \1분 후에 시작됩니다', s)
            s = re.sub('This live event will begin in (\d+) hours', r'이 라이브 이벤트는 \1시간 후에 시작됩니다', s)
            s = re.sub('This live event will begin in (\d+) days', r'이 라이브 이벤트는 \1일 후에 시작됩니다', s)
            s = s.replace('Video unavailable. This video has been removed by the uploader',
                         '비디오를 사용할 수 없습니다. 이 비디오는 업로더에 의해 제거되었습니다') \
                .replace('Unable to extract playlist data', '재생 목록 데이터를 추출할 수 없습니다') \
                .replace('This request has been blocked due to its TLS fingerprint. '
                         'Install a required impersonation dependency if possible, '
                         'or else if you are okay with compromising your security/cookies, '
                         'try replacing "https:" with "http:" in the input URL. '
                         'If you are using a data center IP or VPN/proxy, your IP may be blocked.',
                         '해당 요청은 TLS 지문 때문에 차단되었습니다.\n\n가능하다면 필수 사칭 종속성을 설치하거나, '
                         '보안/쿠키를 손상해도 괜찮다면 입력 URL에서 "https:"를 "http:"로 교체해 보세요. '
                         '데이터 센터 IP 또는 VPN/proxy를 사용하는 경우 IP가 차단될 수 있습니다.')
            s_ = re.sub(r'\n\n가능하다면 필수 사칭 종속성을 설치하거나.*', '', s)
            if 'is not a valid URL' in s_:
                s_ = re.sub("ERROR: \[generic] (.*?) is not a valid URL.+", f"유효하지 않은 URL: {parent.txtURL.GetValue()}", s_)

            elif 'Unsupported URL' in s_:
                s_ = re.sub("ERROR: Unsupported URL: (.*)$", f"지원되지 않는 URL: {parent.txtURL.GetValue()}", s_)

            parent.status.SetLabel(s_)
            parent.gauge.SetValue(0)
            wx.MessageBox(f'{msg}\n\n{s_}', TITLE, style=wx.ICON_ERROR)

            if parent.host == 'vimeo' and '"https:"를 "http:"로 교체해 보세요' in s:
                url = parent.txtURL.GetValue()
                url_ = parent.txtURL.GetValue().replace('https:', 'http:')
                message = f'입력 URL의 "https:"를 "http:"로 교체해서 분석할까요?\n\n' \
                          f'{url}\n=> {url_}'
                with wx.MessageDialog(parent, message, TITLE,
                                      style=wx.YES_NO | wx.ICON_QUESTION) as messageDialog:
                    if messageDialog.ShowModal() == wx.ID_YES:
                        parent.worker = None
                        # parent.cancelled = True
                        try:
                            Popen(f'TASKKILL /F /PID {parent.proc.pid} /T', creationflags=0x08000000)
                        except Exception as e:
                            print(e)

                        parent.txtURL.SetValue(url_)
                        parent.onextract()
                        return

            self.abort()
            return

        if s.startswith('WARNING:'):
            print(s)
            return

        if '[info] Available formats' in s:
            self.table_started = True
            return

        if 'Downloading 1 format(s):' in s:
            return

        if s.startswith('{'):
            parent.formats_data['table'] = self.table_lines
            parent.formats_data['json-string'] = s
            parent.task_done = True
            return

        if s.startswith('['):
            self.table_started = False

        if self.table_started and not s.startswith('['):
            # print(s)
            if not s.startswith('-'):
                # yt-dlp 포맷 출력 테이플의 파일사이즈에 근사값의 의미로 사용된 '≈'이 powershell의 표준출력 과정에서 누락되는 것을 보정하기 위함.
                index = s.find('|')
                index2 = s.find('|', index + 1)
                if not self.pipe_pos:
                    self.pipe_pos += [index, index2]
                if self.pipe_pos:
                    if index2 != self.pipe_pos[1]:
                        s = s[:index + 2] + ' ' * (self.pipe_pos[1] - index2) + s[index + 2:]

                # print(s)
                self.table_lines.append(s)
            return

        s = parent.en2ko(s).replace('\\n', '')

        # print(s)
        parent.status.SetLabel(s)
        parent.progress_count += 1

        parent.gauge.SetValue(parent.progress_count)

    def download_it(self):
        parent = self.parent
        parent.filesize_checked = False
        parent.dn = '[다운로드]'
        fmt = ''
        tempfile = f'{parent.outfile}.part'
        # print(tempfile)
        if tempfile not in parent.tempfiles:
            parent.tempfiles.append(tempfile)

        if parent.cur_video['format'][0] and parent.cur_video['format'][1]:
            fmt = f'{parent.cur_video["format"][0]}+{parent.cur_video["format"][1]}'
        elif parent.cur_video['format'][0]:
            fmt = parent.cur_video['format'][0]

        url = parent.txtURL.GetValue()
        legacyserverconnect = '--legacy-server-connect' if parent.host == 'afreecatv' else ''
        outfile = parent.outfile.encode('utf-16', 'surrogatepass').decode('utf-16')
        cmd = f'powershell & "{YT_DLP_PATH}" -f "{fmt}" "{url}" {legacyserverconnect} -P "{parent.config["download-dir"]}" -o "{outfile}"'.split() + \
              '2>&1 | % ToString | Tee-Object out.txt'.split()
        # print('task =>', parent.task, ', cmd =>', ' '.join(cmd))

        parent.proc = Popen(cmd, stdin=PIPE, stdout=PIPE, stderr=PIPE, creationflags=0x08000000)
        while parent.proc.poll() is None:
            self.checkproc_download()

        if parent.task_done:
            wx.PostEvent(parent, ResultEvent('download-finished'))

    def checkproc_download(self):
        parent = self.parent
        s = str(parent.proc.stdout.readline())
        #print(s)
        if s == "b''":
            return

        s = s.replace("b'", "").replace("\\r\\n'", "")
        s = s.replace('b"', '').replace('\\r\\n"', '')
        print(s)
        if 'getaddrinfo failed' in s or \
            'Unable to download webpage' in s:

            msg = '인터넷 연결상태를 확인해주세요.'

            s = re.sub('.*Unable to download webpage.*', '웹페이지를 다운로드할 수 없습니다.', s)
            s = re.sub('.*getaddrinfo failed.*', 'DNS 조회에 실패했습니다.', s)

            parent.status.SetLabel(s)
            parent.gauge.SetValue(0)
            wx.MessageBox(f'{msg}\n\n{s}', TITLE, style=wx.ICON_ERROR)
            self.abort()
            return

        if s.startswith('WARNING:'):
            if '--paths is ignored since an absolute path is given in output template' in s:
                return

            elif 'Failed to parse XML: not well-formed (invalid token):' in s:
                print(s)
                return

            if 'Inconsistent state of incomplete fragment download' in s:
                s = re.sub('.*Inconsistent state of incomplete fragment download. Restarting from the beginning.*',
                           '조각 다운로드 상태가 일관성이 없는 상태입니다. 처음부터 다시 시작합니다.', s)
                print(s)
                parent.status.SetLabel(s)
                return

        elif s.startswith('[download]'):
            if '100%' in s and parent.dn == '[다운로드]':
                parent.task_done = True

            if s.startswith('[download] Destination:'):
                parent.cur_gauge = parent.gauge.GetValue()
                if parent.L3[parent.row_1st][15] == 'X':
                    if parent.dn == '[다운로드]':
                        parent.dn = '[다운로드-영상]'
                    elif parent.dn == '[다운로드-영상]':
                        parent.dn = '[다운로드-음성]'
                        parent.filesize_checked = False

            frag = ''
            if '(frag' in s:
                result = re.search('\(frag (\d+)/(\d+)\)', s)
                if result:
                    frag = result.group(0).replace('frag', '조각')
                    if result.group(1) != parent.frag_last:
                        parent.frag_last = result.group(1)
                        percent = math.ceil(int(result.group(1)) / int(result.group(2)) * 100)
                        parent.progress_count = parent.cur_gauge + percent

            elif '%' in s:
                result = re.search('(\d+\.\d+)%', s)
                if result:
                    if result.group(1) != parent.percent_last:
                        parent.percent_last = result.group(1)
                        percent = math.ceil(float(result.group(1)))
                        parent.progress_count = parent.cur_gauge + percent

                if not parent.filesize_checked:
                    parent.filesize_checked = True
                    result2 = re.search('of (\s*\d+\.\d+[GMK]iB)', s)
                    if result2:
                        if parent.dn in ['[다운로드]', '[다운로드-영상]']:
                            if result2.group(1) != parent.L3[parent.row_1st][5]:
                                # print('result2.group(1) =>', result2.group(1), 'L3[parent.row_1st][5] =>', parent.L3[parent.row_1st][5])
                                parent.L3[parent.row_1st][5] = result2.group(1)
                                parent.dvlc.SetValue(result2.group(1), parent.row_1st, 1)
                                parent.total_size_in_dvlc_3()
                        elif parent.dn == '[다운로드-음성]':
                            if result2.group(1) != parent.L4[parent.row_2nd][3]:
                                parent.L4[parent.row_2nd][3] = result2.group(1)
                                parent.dvlc_2.SetValue(result2.group(1), parent.row_2nd, 2)
                                parent.total_size_in_dvlc_3()
            else:
                if parent.progress_count == parent.gauge.GetRange():
                    parent.progress_count = 8
                else:
                    parent.progress_count += 1

            parent.gauge.SetValue(parent.progress_count)
            s = s.replace('[download]', f'{parent.dn}') \
                .replace('Destination: ', '저장 파일명: ') \
                .replace('%', f'% {frag}') \
                .replace('of', '    크기:') \
                .replace('at', '    속도:')

            s = re.sub(' ETA .*', '', s)
            parent.status.SetLabel(s)
            # print(s)
            return

        elif s.startswith('[youtube]'):
            if 'Downloading m3u8 information' in s:
                parent.cur_video['m3u8'] = True

        elif s.startswith('[info]'):
            result = re.search(r'Downloading 1 format.*? (\d+)\+(\d+)$', s)
            if result:
                parent.cur_video['format'] = [result.group(1), result.group(2)]
            else:
                result2 = re.search(r'Downloading 1 format.*? (\d+)$', s)
                if result2:
                    if result2.group(1) in ['91', '92', '93', '94', '95', '96', '300', '301']:
                        parent.is_live = True

        elif s.startswith('[Merger]'):
            s = s.replace('[Merger]', '[다중화]') \
                .replace('Merging formats into', '비디오·오디오 다중화 => ')

        elif s.startswith('[FixupM3u8]'):
            s = s.replace('[FixupM3u8]', '[m3u8 손질]') \
                .replace('Fixing MPEG-TS in MP4 container of', 'MP4 컨테이너에서 MPEG-TS 손질 : ')

        elif s.startswith('[VideoRemuxer]'):
            s = s.replace('[VideoRemuxer]', '[리먹싱]')
            if 'Not remuxing media file' in s:
                s = s.replace('Not remuxing media file', '리먹싱 하지 않음.')
            else:
                s = s.replace('Remuxing video from mkv to mp4; Destination', 'mkv를 mp4로 리먹싱 => ')

        elif s.startswith('Deleting original file'):
            self.del_origin_cnt += 1
            if self.del_origin_cnt == 2:
                parent.task_done = True

        parent.progress_count += 1
        parent.gauge.SetValue(parent.progress_count)
        s = parent.en2ko(s).replace('\\n', '') \
            .replace('[generic] ', '') \
            .replace('[redirect] ', '')
        parent.status.SetLabel(s)


class WorkerThread2(Thread):
    def __init__(self, parent):
        Thread.__init__(self)
        self.parent = parent
        parent.progress_count = 0
        parent.gauge.SetValue(parent.progress_count)
        s = f'[리먹싱 준비] {parent.infile}'
        parent.status.SetLabel(s)

    def run(self):
        parent = self.parent
        # parent.cancelled = False
        parent.task_done = False
        cmd = f'powershell & "{FFMPEG}" -y -i'.split() + \
              [f'"{parent.infile}"'] + \
              '-c copy -map 0'.split() + \
              [f'"{parent.outfile}"'] + \
              '2>&1 | % ToString | Tee-Object out.txt'.split()

        # print(' '.join(cmd))
        parent.proc = Popen(cmd, stdin=PIPE, stdout=PIPE, stderr=PIPE, creationflags=0x08000000)
        while parent.proc.poll() is None:
            self.checkproc_remux()

        if parent.task_done:
            wx.PostEvent(parent, ResultEvent('remux-finished'))
        else:
            wx.PostEvent(parent, ResultEvent('remux-cancelled'))

    def abort(self):
        if self.parent.task_done:
            wx.PostEvent(self.parent, ResultEvent('remux-finished'))
        else:
            wx.PostEvent(self.parent, ResultEvent('remux-cancelled'))

        self.raise_exception()

    def get_id(self):
        if hasattr(self, '_thread_id'):
            return self._thread_id

        for t in threading.enumerate():
            if t is self:
                return t.native_id

    def raise_exception(self):
        thread_id = self.get_id()
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id,
                                                         ctypes.py_object(SystemExit))

        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0)
            print('예외 발생 실패')

    def checkproc_remux(self):
        parent = self.parent
        s = str(parent.proc.stdout.readline())
        print(s)
        if s == "b''":
            parent.task_done = True
            self.abort()
            return

        timestamp = ''
        speed = 0.0

        if parent.duration:
            if 'time=' in s:
                t = s.split('time=')[1].split(' bitrate')[0].strip()
                if t:
                    if t.startswith('-'):
                        return

                    timestamp = t

                if s.split('speed=')[1].startswith('N/A'):
                    return

                speed = s.split('speed=')[1].split('x')[0]
        else:
            if 'Duration:' in s:
                s2 = s.split(',')[0].split(': ')[1]
                if s2.startswith('N/A'):
                    return

                parent.duration = s2

            if parent.duration:
                timestamp = '00:00:00'

        if parent.duration and timestamp:
            ws = ' ' * 6
            percent = round((getseconds(timestamp) / getseconds(parent.duration)) * 100)
            if percent < 0:
                return

            if percent > 100:
                percent = 100

            msg = f'{percent}%{ws}{timestamp} / {parent.duration}{ws}{speed}배속'

            try:
                parent.progress_count = percent
                parent.gauge.SetValue(percent)
                parent.status.SetLabel(f'[리먹싱] {msg}')

            except Exception as e:
                print(e)


class WorkerThread3(Thread):
    def __init__(self, parent):
        Thread.__init__(self, None)
        self.parent = parent
        self.done_size_percent = 0
        self.done_size = 0
        self.total_size = 0
        self.size_per_sec = 0

    def run(self):
        parent = self.parent
        # parent.cancelled = False
        parent.task_done = False
        if parent.worker4:
            parent.worker4.abort()

        if parent.task == 'ytdlp':
            parent.set_controls(parent.task)
            parent.gauge.SetRange(10)
            s = '[yt-dlp 업데이트] 준비 중입니다. 잠깐만 기다려주세요..'
            parent.status.SetLabel(s)
            parent.progress_count = 1
            parent.gauge.SetValue(parent.progress_count)
            cmd = f'powershell & "{YT_DLP_PATH}" --update'.split() + \
                  '2>&1 | % ToString | Tee-Object out.txt'.split()
            parent.proc = Popen(cmd, stdin=PIPE, stdout=PIPE, stderr=PIPE, creationflags=0x08000000)
            while parent.proc.poll() is None:
                self.checkproc_update_ytdlp()

        elif parent.task == 'kdownloader':
            parent.set_controls(parent.task)
            parent.gauge.SetRange(102)
            s = f'[{TITLE} 설치파일 다운로드] 준비 중입니다. 잠깐만 기다려주세요..'
            parent.status.SetLabel(s)
            url = ('https://github.com/doriok-lab/k-downloader/releases/download/k-downloader_' +
                   parent.kdownloader_latest_version + '/k-downloader-setup.exe')
            parent.outfile = parent.config["download-dir"] + '\\k-downloader-setup.exe'
            response = requests.get(url, stream=True)
            self.total_size = int(response.headers.get('content-length', 0))
            self.done_size = 0
            done_size_percent_last = -1
            with open(parent.outfile, 'wb') as file:
                try:
                    start_time = time.time()
                    for data in response.iter_content(chunk_size=1024):
                        size = file.write(data)
                        self.done_size += size
                        self.done_size_percent = math.floor(100 * self.done_size / self.total_size)
                        if self.done_size_percent != done_size_percent_last:
                            done_size_percent_last = self.done_size_percent
                            end_time = time.time()
                            elapsed_time = end_time - start_time
                            if elapsed_time == 0:
                                continue

                            start_time = time.time()
                            if round((size / elapsed_time) / 1024, 2) >= 1024:
                                self.size_per_sec = f'{((size / elapsed_time) / 1024) / 1024:8.2f}GiB/s'
                            elif round(size / elapsed_time, 2) >= 1024:
                                self.size_per_sec = f'{(size / elapsed_time) / 1024:8.2f}MiB/s'
                            else:
                                self.size_per_sec = f'{size / elapsed_time:8.2f}KiB/s'

                            self.checkproc_download_kdownloader()

                except Exception as e:
                    print(e)

        #if parent.task_done:
        #    wx.PostEvent(parent, ResultEvent(f'{parent.task}-finished'))

    def abort(self):
        parent = self.parent
        if parent.task_done:
            wx.PostEvent(parent, ResultEvent(f'{parent.task}-finished'))
        else:
            wx.PostEvent(parent, ResultEvent(f'{parent.task}-cancelled'))

        self.raise_exception()

    def get_id(self):
        if hasattr(self, '_thread_id'):
            return self._thread_id

        for t in threading.enumerate():
            if t is self:
                return t.native_id

    def raise_exception(self):
        thread_id = self.get_id()
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id,
                                                         ctypes.py_object(SystemExit))

        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0)
            print('예외 발생 실패')

    def checkproc_update_ytdlp(self):
        parent = self.parent
        s = str(parent.proc.stdout.readline())
        print(s)
        if s == "b''":
            return

        s = s.replace("b'", "").replace("\\r\\n'", "")
        print(s)
        s = s.replace('Current version', '현재 버전') \
            .replace('Latest version', '최신 버전') \
            .replace('Current Build Hash', '현재 빌드 해시') \
            .replace('Updating to stable', '안정 버전으로 업데이트 중...') \
            .replace('Updated yt-dlp to stable', 'yt-dlp 안정 버전으로 업데이트됨.')

        # print(s)
        s = '[yt-dlp] ' + s
        parent.status.SetLabel(s)
        if '안정 버전으로 업데이트됨' in s:
            parent.ytdlp_current_version = parent.ytdlp_latest_version
            parent.gauge.SetValue(10)
            wx.MessageBox(f'yt-dlp 업데이트 완료\n\n업데이트 버전: {parent.ytdlp_latest_version}', TITLE, style=wx.ICON_INFORMATION)
            wx.PostEvent(self.parent, ResultEvent('ytdlp-finished'))
            return

        if '현재 버전' in s or '최신 버전' in s:
            time.sleep(1)

        parent.progress_count += 1
        parent.gauge.SetValue(parent.progress_count)

    def checkproc_download_kdownloader(self):
        parent = self.parent
        s = (f'[K-Downloader 설치파일 다운로드] {self.done_size_percent:4}%      '
             f'크기: {self.total_size/1048576:7.2f}MiB      속도: {self.size_per_sec}')
        parent.status.SetLabel(s)
        parent.gauge.SetValue(self.done_size_percent)
        if self.done_size_percent >= 100:
            parent.task_done = True
            self.abort()
            return


class WorkerThread4(Thread):
    def __init__(self, parent, arg):
        Thread.__init__(self)
        self.parent = parent
        self.arg = arg
        self.checked = False

    def run(self):
        parent = self.parent
        parent.set_controls(parent.task)
        if self.arg == 'ytdlp':
            url = 'https://github.com/yt-dlp/yt-dlp/blob/master/yt_dlp/version.py'
            text = requests.get(url).text
            result = re.search(r"__version__ = '(.*?)'", text)
            if result:
                parent.ytdlp_latest_version = result.group(1).strip()
                if not parent.update_notify_ytdlp:
                    if parent.ytdlp_latest_version != parent.ytdlp_current_version:
                        message = f'yt-dlp 최신 버전이 있습니다. 업데이트할까요?\n\n' \
                                  f'현재 버전: {parent.ytdlp_current_version}\n\n최신 버전: {parent.ytdlp_latest_version}'

                        with wx.MessageDialog(parent, message, TITLE,
                                              style=wx.YES_NO | wx.ICON_QUESTION) as messageDialog:
                            if messageDialog.ShowModal() == wx.ID_YES:
                                parent.task = 'ytdlp'
                                parent.worker3 = WorkerThread3(parent)
                                parent.worker3.daemon = True
                                parent.worker3.start()

        elif self.arg == 'kdownloader':
            url = 'https://github.com/doriok-lab/k-downloader/blob/main/version.py'
            text = requests.get(url).text
            result = re.search(r"__version__ = '(.*?)'", text)
            if result:
                self.checked = True
                parent.kdownloader_latest_version = result.group(1).strip()
                if not parent.update_notify_kdownloader:
                    if parent.kdownloader_latest_version != VERSION:
                        self.abort()

    def abort(self):
        if self.parent.task == 'checkversion':
            if self.checked:
                wx.PostEvent(self.parent, ResultEvent(f'{self.parent.task}-finished'))
            else:
                wx.PostEvent(self.parent, ResultEvent(f'{self.parent.task}-cancelled'))

        self.raise_exception()

    def raise_exception(self):
        thread_id = self.get_id()
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id,
                                                         ctypes.py_object(SystemExit))
        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0)
            print('예외 발생 실패')

    def get_id(self):
        if hasattr(self, '_thread_id'):
            return self._thread_id

        for t in threading.enumerate():
            if t is self:
                return t.native_id


class VideoDownloader(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title=TITLE, size=(880, -1),
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))

        self.duration = ''
        self.wtxt = ''
        self.task = ''
        self.host = ''
        self.infile = ''
        self.outfile = ''
        self.err_msg = ''
        self.frag_last = ''
        self.download_error = ''
        self.dn = ''
        self.ext_origin = ''
        self.ytdlp_current_version = ''
        self.ytdlp_latest_version = ''
        self.kdownloader_latest_version = ''
        self.percent_last = -1
        self.progress_count = 0
        self.row_1st = -1
        self.row_2nd = -1
        self.cur_dvlc = 1
        self.cur_gauge = 0
        self.tempfiles = []
        self.lines = []
        self.pids_explorer_existing = []
        self.cur_video = {}
        self.config = {}
        self.formats_data = {}
        self.control_state = {}
        self.pos = {}
        self.last_sel = (-1, -1)
        self.no_ffmpeg = False
        self.last_check = False
        self.last_radio = False
        self.cbRemux_visible = False
        self.is_live = False
        self.update_notify_ytdlp = False
        self.update_notify_kdownloader = False
        self.filesize_checked = False
        self.task_done = False
        self.worker = None
        self.worker2 = None
        self.worker3 = None
        self.worker4 = None
        self.worker5 = None
        self.wsh = None
        self.dlg = None
        self.proc = None
        self.L3 = None
        self.L3_2 = None
        self.L4 = None

        try:
            with open('config.pickle', 'rb') as f:
                self.config = pickle.load(f)
                if 'remuxing' not in self.config:
                    self.config["remuxing"] = False

                if 'download-dir' not in self.config:
                    self.config["download-dir"] = os.path.expanduser('~') + '\\Downloads'

        except Exception as e:
            print(e)
            self.config["remuxing"] = False
            self.config["download-dir"] = os.path.expanduser('~') + '\\Downloads'

        # 메뉴
        self.menuBar = wx.MenuBar()
        self.menu1 = wx.Menu()
        self.menu1.Append(101, '저장폴더 열기...')
        self.menu1.AppendSeparator()
        self.menu1.Append(111, '리먹싱(=>MP4)...')
        self.menu1.AppendSeparator()
        self.menu1.Append(109, '종료')
        self.menuBar.Append(self.menu1, '  파일  ')
        self.menu2 = wx.Menu()
        self.menu2.Append(203, 'yt-dlp 업데이트')
        self.menu2.AppendSeparator()
        self.menu2.Append(205, '업데이트')
        self.menu2.Append(204, '정보')
        self.menuBar.Append(self.menu2, '  도움말  ')
        self.SetMenuBar(self.menuBar)

        # 컨트롤
        pn = wx.Panel(self)
        pn.SetDoubleBuffered(True)
        # font = wx.Font(8, wx.MODERN, wx.NORMAL, wx.NORMAL, faceName='맑은 고딕')
        # pn.SetFont(font)
        # pn.Font.SetPointSize(11)
        st = wx.StaticText(pn, -1, "동영상 URL", size=(70, -1))
        self.txtURL = txtURL = wx.TextCtrl(pn, -1, '')
        self.btnPaste = btnPaste = wx.Button(pn, -1, '⇋', size=(25, 25))
        btnPaste.SetToolTip('매크로 실행 => "동영상URL 복붙" 및 "정보 추출"')
        btnPaste.SetBackgroundColour((255, 255, 255))
        self.btnExtract = btnExtract = wx.Button(pn, -1, '정보 추출')
        self.status = status = wx.StaticText(pn, -1, '', size=(-1, 30))
        self.btnOpenDir = btnOpenDir = wx.Button(pn, -1, '폴더 열기')
        self.btnOpen = btnOpen = wx.Button(pn, -1, '재생')
        self.gauge = gauge = wx.Gauge(pn, -1, 7)
        self.btnAbort = btnAbort = wx.Button(pn, -1, '중지')
        btnAbort.Disable()
        self.dvlc = dvlc = wx.dataview.DataViewListCtrl(pn,
                                                        size=(280, 332), style=wx.BORDER_NONE)
        col_labels = [('포맷', 185), ('파일 크기', 75)]
        for x, y in col_labels:
            dvlc.AppendTextColumn(x, width=y)

        self.dvlc_2 = dvlc_2 = wx.dataview.DataViewListCtrl(pn,
                                                            size=(250, 332), style=wx.BORDER_NONE)
        col_labels2 = [('확장자', 65), ('오디오 BR', 70), ('파일 크기', 75), ('비고', 30)]
        for x, y in col_labels2:
            dvlc_2.AppendTextColumn(x, width=y)

        self.dvlc_3 = dvlc_3 = wx.dataview.DataViewListCtrl(pn,
                                                            size=(250, 332), style=wx.BORDER_NONE)
        col_labels3 = [('항목', 135), ('값', 95)]
        for x, y in col_labels3:
            dvlc_3.AppendTextColumn(x, width=y)

        dvlc_3.SetBackgroundColour('#efefef')
        dvlc_3.Disable()

        btnhelp = wx.Button(pn, -1, '?', size=(22, -1))
        btnhelp.SetBackgroundColour((255, 255, 255))
        btnhelp.SetToolTip('"포맷 선택" 도움말')

        self.btnPreview = btnPreview = wx.Button(pn, -1, '미리보기')
        btnPreview.Disable()

        btnhelp2 = wx.Button(pn, -1, '?', size=(22, -1))
        btnhelp2.SetBackgroundColour((255, 255, 255))
        btnhelp2.SetToolTip('"오디오 추가" 도움말')
        self.btnUnselect = btnUnselect = wx.Button(pn, -1, '선택해제')
        btnUnselect.Disable()

        self.btnPreview2 = btnPreview2 = wx.Button(pn, -1, '미리듣기')
        btnPreview2.Disable()

        self.cbRemux = wx.CheckBox(pn, -1, "후처리:리먹싱(=>MP4)", size=(-1, 22))
        self.cbRemux.SetValue(self.config["remuxing"])
        self.cbRemux.SetToolTip('다운로드된 파일이 MP4가 아닌 경우(mkv, webm, m3u8 등) MP4로 리먹싱함.')
        self.cbRemux.Disable()

        self.btnHelp3 = wx.Button(pn, -1, '?', size=(22, -1))
        self.btnHelp3.SetBackgroundColour((255, 255, 255))
        self.btnHelp3.SetToolTip('"후처리:리먹싱(=>MP4)" 도움말')

        self.btnDownload = wx.Button(pn, -1, '다운로드')
        self.btnDownload.Disable()
        self.btnDownloadDir = wx.Button(pn, -1, '⚙', size=(22, 22))
        font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_MAX)
        self.btnDownloadDir.SetFont(font)
        self.btnDownloadDir.SetBackgroundColour((255, 255, 255))
        self.btnDownloadDir.SetToolTip('다운로드 파일을 저장할 폴더를 지정합니다.')

        self.SetIcon(wx.Icon("data/k-downloader.ico"))

        # 레이아웃
        box = wx.StaticBox(pn, -1, '')
        bsizer = wx.StaticBoxSizer(box, wx.HORIZONTAL)
        bsizer.Add(st, 0, wx.LEFT | wx.TOP, 10)
        bsizer.Add(txtURL, 1, wx.LEFT | wx.TOP | wx.BOTTOM, 5)
        bsizer.Add(btnPaste, 0, wx.TOP, 4)
        bsizer.Add(btnExtract, 0, wx.LEFT | wx.TOP | wx.BOTTOM, 5)
        inner = wx.BoxSizer()
        inner.Add(bsizer, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        self.inner2_1 = inner2_1 = wx.BoxSizer(wx.HORIZONTAL)
        inner2_1.Add(status, 1, wx.EXPAND | wx.TOP, 5)
        inner2_1.Add(btnOpenDir, 0, wx.LEFT, 5)
        inner2_1.Add(btnOpen, 0, wx.LEFT, 5)

        self.inner2_2 = inner2_2 = wx.BoxSizer(wx.HORIZONTAL)
        inner2_2.Add(gauge, 1, wx.TOP, 5)
        inner2_2.Add(btnAbort, 0, wx.LEFT, 5)

        self.inner2 = inner2 = wx.BoxSizer(wx.VERTICAL)
        inner2.Add(inner2_2, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)
        inner2.Add(inner2_1, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)
        box = wx.StaticBox(pn, -1, '다운로드')
        self.bsizer3 = bsizer3 = wx.StaticBoxSizer(box, wx.HORIZONTAL)
        bsizer3.Add(dvlc_3, 0, wx.ALL, 5)
        inner3_1 = wx.BoxSizer(wx.HORIZONTAL)
        inner3_1.Add(btnhelp, 0, wx.RIGHT, 10)
        inner3_1.Add((1, -1), 1, wx.RIGHT, 10)
        inner3_1.Add(btnPreview, 0, wx.RIGHT, 10)

        box = wx.StaticBox(pn, -1, '포맷 선택')
        bsizer = wx.StaticBoxSizer(box, wx.VERTICAL)
        bsizer.Add(dvlc, 0, wx.ALL, 5)
        bsizer.Add(inner3_1, 0, wx.EXPAND | wx.ALL, 5)

        inner3_2 = wx.BoxSizer(wx.HORIZONTAL)
        inner3_2.Add(btnhelp2, 0, wx.RIGHT, 10)
        inner3_2.Add((1, -1), 1, wx.RIGHT, 10)
        inner3_2.Add(btnUnselect, 0, wx.RIGHT, 10)
        inner3_2.Add(btnPreview2, 0, wx.RIGHT, 10)

        box = wx.StaticBox(pn, -1, '오디오 추가')
        self.bsizer2 = bsizer2 = wx.StaticBoxSizer(box, wx.VERTICAL)
        bsizer2.Add(dvlc_2, 0, wx.ALL, 5)
        bsizer2.Add(inner3_2, 0, wx.EXPAND | wx.ALL, 5)

        self.inner3 = inner3 = wx.BoxSizer(wx.HORIZONTAL)
        inner3.Add(bsizer, 0)
        inner3.Add(bsizer2, 0, wx.LEFT, 10)
        inner3.Add(bsizer3, 0, wx.LEFT, 10)

        self.inner4 = inner4 = wx.BoxSizer(wx.HORIZONTAL)
        inner4.Add((1, -1), 1, wx.RIGHT, 10)
        inner4.Add(self.cbRemux, 0)
        inner4.Add(self.btnHelp3, 0, wx.RIGHT, 10)
        inner4.Add((1, -1), 0, wx.RIGHT, 10)
        inner4.Add(self.btnDownload, 0, wx.RIGHT | wx.BOTTOM, 10)
        inner4.Add(self.btnDownloadDir, 0, wx.RIGHT | wx.BOTTOM, 10)

        self.sizer = sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(inner, 0, wx.EXPAND | wx.TOP | wx.RIGHT | wx.LEFT, 5)
        sizer.Add(inner2, 0, wx.EXPAND | wx.RIGHT | wx.LEFT, 5)
        sizer.Add(inner3, 0, wx.EXPAND | wx.TOP | wx.RIGHT | wx.LEFT, 10)
        sizer.Add(inner4, 0, wx.EXPAND | wx.ALL, 10)

        pn.SetSizer(sizer)
        sizer.SetSizeHints(self)
        self.Center()

        inner2_1.Hide(btnOpenDir)
        inner2_1.Hide(btnOpen)
        sizer.Hide(inner3)
        sizer.Hide(inner4)

        # 이벤트
        self.Bind(wx.EVT_MENU, self.onopen_dir, id=101)
        self.Bind(wx.EVT_MENU, self.onremux, id=111)
        self.Bind(wx.EVT_MENU, self.onclose, id=109)
        self.Bind(wx.EVT_MENU, self.onupdate_ytdlp, id=203)
        self.Bind(wx.EVT_MENU, self.onupdate_kdownloader, id=205)
        self.Bind(wx.EVT_MENU, self.onabout, id=204)
        self.Bind(wx.EVT_CLOSE, self.onwindow_close)
        txtURL.Bind(wx.EVT_TEXT, self.oncheck_url)
        btnPaste.Bind(wx.EVT_BUTTON, self.onpaste_it)
        btnExtract.Bind(wx.EVT_BUTTON, self.onextract)
        btnOpenDir.Bind(wx.EVT_BUTTON, self.onopen_dir2)
        btnOpen.Bind(wx.EVT_BUTTON, self.onopen_file)
        btnAbort.Bind(wx.EVT_BUTTON, self.onabort)
        dvlc.Bind(wx.dataview.EVT_DATAVIEW_SELECTION_CHANGED, self.onsel_format)
        dvlc_2.Bind(wx.dataview.EVT_DATAVIEW_SELECTION_CHANGED, self.onsel_format_2)
        btnhelp.Bind(wx.EVT_BUTTON, self.onhelp)
        btnPreview.Bind(wx.EVT_BUTTON, self.onpreview)
        btnhelp2.Bind(wx.EVT_BUTTON, self.onhelp2)
        btnUnselect.Bind(wx.EVT_BUTTON, self.onunselect_it)
        btnPreview2.Bind(wx.EVT_BUTTON, self.onpreview2)
        self.cbRemux.Bind(wx.EVT_CHECKBOX, self.oncheck_remux)
        self.btnHelp3.Bind(wx.EVT_BUTTON, self.onhelp_it3)
        self.btnDownload.Bind(wx.EVT_BUTTON, self.ondownload)
        self.btnDownloadDir.Bind(wx.EVT_BUTTON, self.ondownload_dir)
        self.Bind(wx.EVT_CHAR_HOOK, self.onkey)
        evt_result(self, self.onresult)

        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'explorer.exe':
                self.pids_explorer_existing.append(proc.info['pid'])

        self.ytdlp_current_version = get_version_number(YT_DLP_PATH)
        wx.CallLater(1, self.check_version_latest, 'ytdlp')
        wx.CallLater(1000, self.check_version_latest, 'kdownloader')

    def onkey(self, evt):
        obj = wx.Window.FindFocus()
        if evt.GetKeyCode() == wx.WXK_RETURN:
            if self.txtURL.HasFocus():
                if self.worker or self.worker2 or self.worker3 or self.worker5:
                    self.onabort(evt)
                else:
                    self.onextract()
            elif obj.ClassName in ['wxDataViewMainWindow', 'wxCheckBox']:
                self.ondownload()
            else:
                evt.Skip()
        else:
            evt.Skip()

    def check_version_latest(self, arg=None):
        self.task = 'checkversion'
        self.worker4 = WorkerThread4(self, arg)
        self.worker4.daemon = True
        self.worker4.start()

    def check_integral(self, n):
        result = re.search(r'^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$', n)
        if not result:
            result2 = re.search(r'^([0-5]\d):([0-5]\d)$', n)
            if not result2:
                result3 = re.search(r'^\d+$', n)
                if not result3:
                    self.err_msg = "시간은 '시:분:초' 형식으로 입력해주세요." \
                                   "\n이때 시, 분, 초는 각각 2자리의 정수입니다. " \
                                   "\n(예시) '01:23:45', '01:23' " \
                                   "\n\n한편, '시:분:초' 형식 대신에 그냥 초로 환산하여 입력할 수도 있습니다." \
                                   "\n즉, '01:23' 대신 '83'도 됩니다."
                    return False

        return True

    def check_decimal(self, n):
        result = re.search(r'^(\d+)$', n)
        if not result:
            self.err_msg = '소수점 아래는 숫자 외의 문자를 사용할 수 없습니다.'
            return False

        return True

    def check_range(self, s):
        s = re.sub(' ', '', s)
        if '.' in s:
            if s.count('.') > 1:
                self.err_msg = '소수점은 초를 표시하는 부분에만 사용할 수 있습니다.'
                return False
            else:
                integral, decimal = s.split('.')
                if not self.check_integral(integral):
                    return False

                if not self.check_decimal(decimal):
                    return False
        else:
            if not self.check_integral(s):
                return False

        return True

    def oncheck_url(self, evt):
        if self.txtURL.GetValue().strip() == '':
            self.btnPaste.Disable()
            self.btnExtract.Disable()
        else:
            self.btnPaste.Enable()
            self.btnExtract.Enable()

    def onabort(self, evt=None):
        self.status.SetLabel('취소 중...')
        if self.worker:
            self.worker.abort(evt)

        elif self.worker2:
            self.worker2.abort()

        elif self.worker3:
            self.worker3.abort()

        elif self.worker4:
            self.worker4.abort()

        elif self.worker5:
            self.worker5.abort()


    def onpaste_it(self, evt):
        text = clipboard.paste()
        self.txtURL.SetValue(text)
        self.btnExtract.SetFocus()
        self.onextract()

    def onextract(self, evt=None):
        if not self.txtURL.GetValue():
            wx.MessageBox('동영상 URL을 입력하세요.\n\n ', TITLE, style=wx.ICON_WARNING)
            self.txtURL.SetFocus()
            return

        self.dvlc.DeleteAllItems()
        self.dvlc_2.DeleteAllItems()
        self.dvlc_3.DeleteAllItems()
        self.row_1st = -1
        self.row_2nd = -1
        self.last_sel = (-1, -1)

        self.task = 'extract'
        self.wtxt = ''
        self.L3 = None
        self.L3_2 = None
        self.L4 = None
        self.gauge.SetRange(20 if 'vimeo.com/' in self.txtURL.GetValue() else 10)
        self.cur_gauge = 0
        self.set_controls(self.task)
        self.worker = WorkerThread(self)
        self.worker.daemon = True
        self.worker.start()

    def onresult(self, evt):
        if self.task in ['extract', 'download']:
            self.worker = None

        elif self.task == 'remux':
            self.worker2 = None

        elif self.task == 'ytdlp':
            self.worker3 = None

        elif self.task == 'checkversion':
            self.worker4 = None

        elif self.task == 'kdownloader':
            self.worker5 = None

        self.btnAbort.Disable()
        self.btnExtract.Enable()
        self.dvlc.Enable()
        self.dvlc_2.Enable()

        if evt.data:
            self.restore_controls(self.task)
            if evt.data in ['extract-cancelled', 'extract-cancelled2']:
                self.gauge.SetValue(0)
                msg = '정보 추출이 취소되었습니다.'
                self.status.SetLabel(msg)
                arg = 1 if evt.data == 'extract-cancelled2' else None
                self.killtask(msg, arg)
                self.sizer.Hide(self.inner3)
                self.sizer.Hide(self.inner4)
                self.txtURL.SelectAll()

            elif evt.data in ['download-cancelled', 'download-cancelled2']:
                self.gauge.SetValue(0)
                msg = '다운로드가 취소되었습니다.'
                self.status.SetLabel(msg)
                arg = 1 if evt.data == 'download-cancelled2' else None
                self.killtask(msg, arg)

            elif evt.data == 'remux-cancelled':
                self.gauge.SetValue(0)
                msg = '리먹싱이 취소되었습니다.'
                self.status.SetLabel(msg)
                self.killtask(msg)

            elif evt.data in ['ytdlp-cancelled', 'kdownloader-cancelled']:
                self.gauge.SetValue(0)
                msg = 'yt-dlp 업데이트가 취소되었습니다.' if self.task == 'ytdlp' \
                    else f'{TITLE} 설치파일 다운로드가 취소되었습니다.'
                self.status.SetLabel(msg)
                self.killtask(msg)

            elif evt.data == 'extract-finished':
                self.display_formats('table')

            elif evt.data == 'download-finished':
                if self.no_ffmpeg:
                    self.status.SetLabel(f'[비디오·오디오 합치기] FFmpeg 설치되어 있지 않음. 합치기 실패.')
                    self.gauge.SetValue(0)
                else:
                    self.gauge.SetValue(self.gauge.GetRange())
                    filename = os.path.split(self.outfile)[1]
                    self.status.SetLabel(f'[다운로드 완료] {filename}')
                    self.btnOpenDir.Show()
                    self.btnOpen.Show()
                    self.inner2.Layout()
                    if self.cbRemux_visible and self.config["remuxing"]:
                        wx.CallLater(1, self.remux_it, self.outfile)
                    else:
                        self.onopen_dir2()

            elif evt.data == 'remux-finished':
                dirname_basename = os.path.splitext(self.outfile)[0]
                file_origin = f'{dirname_basename}.{self.ext_origin}'
                # print('file_origin =>', file_origin)
                if os.path.isfile(file_origin):
                    os.remove(file_origin)

                self.gauge.SetValue(self.gauge.GetRange())
                filename = os.path.split(self.outfile)[1]
                self.status.SetLabel(f'[리먹싱 완료] {filename}')

                # wx.MessageBox(f'리먹싱 완료\n\n{filename}', TITLE, style=wx.ICON_INFORMATION)
                self.btnOpenDir.Show()
                self.btnOpen.Show()
                self.inner2.Layout()
                self.onopen_dir2()

            elif evt.data == 'checkversion-finished':
                self.onupdate_kdownloader()
                pass

            elif evt.data == 'checkversion-cancelled':
                pass

            elif evt.data == 'ytdlp-finished':
                pass

            elif evt.data == 'kdownloader-finished':
                self.btnOpen.Enable(self.task != 'kdownloader')
                if self.task == 'kdownloader':
                    self.gauge.SetValue(self.gauge.GetRange())
                    filename = os.path.split(self.outfile)[1]
                    self.status.SetLabel(f'[{TITLE} 설치파일 다운로드 완료] {filename}')
                    self.btnOpenDir.Show()
                    self.btnOpen.Show()
                    self.inner2.Layout()
                    message = f'업데이트를 진행하려면 일단 프로그램을 닫은 후 설치파일을 실행해야 합니다. 계속 할까요?\n\n' \
                              f'설치파일: {self.outfile}'

                    with wx.MessageDialog(self, message, f'{TITLE} 업데이트',
                                          style=wx.YES_NO | wx.ICON_QUESTION) as messageDialog:
                        if messageDialog.ShowModal() == wx.ID_YES:
                            self.Close()
                            self.onopen_dir2()

    def killtask(self, message, arg=None):
        if self.proc:
            Popen(f'TASKKILL /F /PID {self.proc.pid} /T', creationflags=0x08000000)

        if not arg:
            wx.MessageBox(f'{message}\n\n ', TITLE)

    def display_formats(self, arg):
        self.dvlc.DeleteAllItems()
        self.inner3.Hide(self.bsizer2)
        self.inner3.Hide(self.bsizer3)
        self.cbRemux.Hide()
        self.cbRemux_visible = False
        self.btnHelp3.Hide()
        if arg == 'table':
            self.cur_video["id"], self.cur_video["title"], self.cur_video["duration"], self.cur_video["is_live"] = \
                self.get_video_info(self.formats_data['json-string'])

            if self.cur_video["is_live"]:
                self.gauge.SetValue(0)
                msg = f'[{self.host}] {self.cur_video["id"]} : 라이브 스트리밍 중...'
                self.status.SetLabel(msg)
                self.restore_controls('extract')
                self.sizer.Hide(self.inner3)
                self.sizer.Hide(self.inner4)
                message_ = f'라이브 스트리밍 중...\n\n[{self.host}] {self.cur_video["id"]} :\n{self.cur_video["title"]}'
                wx.MessageBox(message_, TITLE, wx.ICON_EXCLAMATION | wx.OK)
                return

            self.lines = []
            self.pos = {}
            self.extract_data()

            mylist = []
            print(
                '0=>ID, 1=>EXT, 2=>RESOLUTION, 3=>FPS, 4=>CH, 5=>FILESIZE, 6=>TBR, 7=>PROTO, 8=>VCODEC, 9=>VBR, 10=>ACODEC, 11=>ABR, 12=>ASR, 13=>MORE INFO, 14=>VIDEO(O/X), 15=>AUDIO(O/X), 16=>url')

            for line in self.lines:
                if line.startswith('ID'):
                    continue

                l = []
                l2 = []
                for key in self.pos:
                    begin = self.pos[key][0]
                    end = self.pos[key][1]
                    if begin == -1 or end == -1:
                        val = ''
                    else:
                        val = line[begin:end].rstrip()

                    l.append(val)

                if l[2] == 'audio only':
                    l2.append('X')
                else:
                    l2.append('O')

                if l[10] == 'video only':
                    l2.append('X')
                else:
                    l2.append('O')

                url = self.get_video_url(self.formats_data['json-string'], l[0])
                l2.append(url)
                l += l2
                mylist.append(l)

            self.L3 = L3 = []
            self.L3_2 = L3_2 = []

            #  포맷 목록
            mylist2 = []
            for l in mylist:
                if l[1] in ['mp4', 'webm'] and l[7] in ['http', 'https'] and l[
                    14] == 'O':  # 1=>ext, 7=>proto, 14=>video(O/X)
                    if self.host == 'facebook':
                        if l[2] == 'unknown':
                            l[2] = l[0]
                    elif self.host == 'naver':
                        # 파일 크기가 모두 동일하게 나옴 => 재생시간, 총비트전송률을 고려하여 파일 크기 계산
                        tbr = int(l[6][:-1]) * 1000
                        size_byte = tbr / 8 * int(self.cur_video["duration"])
                        unit = ''
                        filesize = 0
                        if size_byte >= 1073741824:
                            unit = 'GiB'
                            filesize = f'{round(size_byte / 1073741824, 2):.2f}'
                        elif size_byte >= 1048576:
                            unit = 'MiB'
                            filesize = f'{round(size_byte / 1048576, 2):.2f}'
                        elif size_byte >= 1024:
                            unit = 'KiB'
                            filesize = f'{round(size_byte / 1024, 2):.2f}'

                        l[5] = f'{filesize:>7s}{unit}'

                    print("v protocol:http or https", l)
                    mylist2.append(l)

            for l in mylist:
                if self.host == 'youtube':
                    if l[1] in ['mp4', 'webm'] and l[7] == 'm3u8' and l[
                        14] == 'O':  # 1=>ext, 7=>proto, 14=>video(O/X), 15=>audio(O/X)
                        print("v protocol:m3u8", l)
                        mylist2.append(l)

                elif self.host == 'vimeo':
                    if 'fastly' in l[0]:
                        if l[1] in ['mp4', 'webm'] and l[7] == 'm3u8' and l[
                            14] == 'O':  # 1=>ext, 7=>proto, 14=>video(O/X)
                            print("v protocol:m3u8", l)
                            mylist2.append(l)
                else:
                    if l[1] in ['mp4', 'webm'] and l[7] == 'm3u8' and l[
                        14] == 'O':  # 1=>ext, 7=>proto, 14=>video(O/X)
                        print("v protocol:m3u8", l)
                        mylist2.append(l)

            for l in mylist:
                if l[1] in ['mp4', 'webm'] and l[7] == 'dash' and l[14] == 'O':  # 1=>ext, 7=>proto, 14=>video(O/X)
                    print("v protocol:dash", l)
                    mylist2.append(l)

            if mylist2:
                mylist2.sort(key=lambda col: (
                2 if col[7] in ['http', 'https'] else (1 if col[7] == 'm3u8' else 0), 1 if col[1] == 'mp4' else 0,
                int(col[2].split('x')[1]) if 'x' in col[2] else (0 if col[2] == 'sd' else 1),
                int(col[2].split('x')[0]) if 'x' in col[2] else (0 if col[2] == 'sd' else 1),
                col[3], col[6], col[5]), reverse=True)  # 1=> exp, 2=>resolution, 3=>fps, 5=>filesize, 6=>tbr, 7=> proto
                L3 += mylist2

            no_audio = False
            for l in L3:
                if l[15] == 'X':
                    no_audio = True
                    break

            # 오디오 추가
            if no_audio:
                mylist2 = []
                for l in mylist:
                    if l[1] in ['m4a', 'webm'] and l[7] in ['http', 'https'] and l[
                        15] == 'O':  # 1=>ext, 7=>proto, 15=>audio(O/X)
                        print("a protocol:http or https", l)
                        mylist2.append(l)

                # if not mylist2:
                if True:
                    for l in mylist:
                        if self.host == 'vimeo':
                            if 'fastly' in l[0]:
                                if l[1] == 'mp4' and l[7] == 'm3u8' and l[
                                    15] == 'O':  # 1=>ext, 7=>proto, 11=>abr, 15=>audio(O/X)
                                    if 'high' in l[0]:
                                        l[11] = '253k'
                                    else:
                                        l[11] = '125k'

                                    abr = 253375 if l[11] == '253k' else (
                                        125375 if l[11] == '125k' else int(l[11][:-1]) * 1000)
                                    size_byte = abr / 8 * int(self.cur_video["duration"])
                                    unit = ''
                                    filesize = 0
                                    if size_byte >= 1073741824:
                                        unit = 'GiB'
                                        filesize = f'{round(size_byte / 1073741824, 2):.2f}'
                                    elif size_byte >= 1048576:
                                        unit = 'MiB'
                                        filesize = f'{round(size_byte / 1048576, 2):.2f}'
                                    elif size_byte >= 1024:
                                        unit = 'KiB'
                                        filesize = f'{round(size_byte / 1024, 2):.2f}'

                                    l[5] = f'{filesize:>7s}{unit}'
                                    print("a protocol:m3u8", l)
                                    mylist2.append(l)
                        else:
                            if l[1] == 'mp4' and l[7] == 'm3u8' and l[11] != '' and l[
                                15] == 'O':  # 1=>ext, 7=>proto, 11=>abr,15=>audio(O/X)
                                print("a protocol:m3u8", l)
                                mylist2.append(l)

                for l in mylist:
                    if l[1] == 'm4a' and l[7] == 'dash' and l[15] == 'O':  # 1=>ext, 7=>proto, 15=>audio(O/X)
                        print("a protocol:dash", l)
                        mylist2.append(l)

                if mylist2:
                    mylist2.sort(key=lambda col: (2 if col[7] in ['http', 'https'] else (1 if col[7] == 'm3u8' else 0),
                                             1 if col[1] in ['m4a', 'mp4'] else 0, col[11], col[5],
                                             1 if 'DRC' in col[13] else 0),
                            reverse=True)  # 1=>ext, 5=>filesize, 11=>abr, 13=>more info
                    L3_2 += mylist2

            for l in L3:
                l = [f'{l[2]} ( {l[1]}, {l[6].strip()} )', l[5]] if l[6] else [f'{l[2]} ( {l[1]} )', l[
                    5]]  # 1=>ext, 2=>resolution, 5=>filesize, 6=>tbr
                self.dvlc.AppendItem(l)

            self.dvlc.Layout()
            msg = f'[정보 추출 완료] {self.cur_video["id"]}: {self.cur_video["title"]}'
            self.status.SetLabel(msg)
            self.gauge.SetValue(self.gauge.GetRange())

            self.L4 = []
            for l in L3_2:
                self.L4.append([l[0], l[1], l[4], l[5], l[10], l[11], l[12], l[16], l[
                    7]])  # 0=>id, 1=>ext, 4=>ch, 5=>filesize, 10=>acodec, 11=>abr, 12=>asr, 16=>url, 7=>proto
                drc = 'DRC' if l[0].endswith('-drc') else ''
                self.dvlc_2.AppendItem([l[1], l[11], l[5], drc])  # 1=>ext, 5=>filesize, 11=>abr

            self.dvlc_init()

        self.sizer.Show(self.inner3)
        self.sizer.Show(self.inner4)
        self.sizer.Layout()

    def get_video_info(self, data):
        # print(data)
        index = data.find('"id":')
        index2 = data.find('"', index + 6)
        index3 = data.find('"', index2 + 1)
        video_id = data[index2 + 1:index3]
        index4 = data.find('"title":')
        index5 = data.find(' "', index4 + 8)
        index6 = data.find('",', index5 + 1)
        video_title = data[index5 + 2:index6].replace("|", "#") \
            .encode().decode('unicode-escape').encode().decode('unicode-escape')

        video_duration = ''
        if self.host in ['naver', 'vimeo']:
            index7 = data.find('"duration":', index6 + 1)
            index8 = data.find(',', index7 + 11)
            video_duration = data[index7 + 12:index8]

        if self.host == 'facebook':
            index9 = data.find('"format_id":')
            index10 = data.find(',', index9 + 12)
            is_live = True if 'dash-lp-md' in data[index9 + 13:index10] else False
        else:
            index9 = data.find('"is_live":')
            index10 = data.find(',', index9 + 10)
            is_live = True if data[index9 + 11:index10] == 'true' else False

        return video_id, video_title, video_duration, is_live

    @staticmethod
    def get_video_url(data, format_id):
        needle = f'"format_id": "{format_id}"'
        index = data.find(needle)
        index2 = data.find('"url":', index + 18)
        index3 = data.find('"', index2 + 7)
        index4 = data.find('"', index3 + 1)
        url = data[index3 + 1:index4]
        return url

    def extract_data(self):
        self.lines = self.formats_data['table']

        self.pos = {}
        keys = ['ID', 'EXT', 'RESOLUTION', 'FPS', 'CH', 'FILESIZE', 'TBR', 'PROTO', 'VCODEC', 'VBR', 'ACODEC', 'ABR',
                'ASR', 'MORE INFO']
        for key in keys:
            self.pos[key] = [-1, -1]

        header = self.lines[0]
        left_aligns = ['ID', 'EXT', 'RESOLUTION', 'PROTO', 'VCODEC', 'ACODEC', 'MORE INFO']
        for key in left_aligns:
            idx = header.find(key)
            self.pos[key][0] = idx

        if self.pos['MORE INFO'][0] != -1:
            self.pos['MORE INFO'][1] = 1000

        right_aligns = ['ASR', 'ABR', 'VBR', 'TBR', 'FILESIZE', 'CH', 'FPS']
        for key in right_aligns:
            if key in header:
                idx = header.find(key)
                if key in ['CH', 'FPS']:
                    self.pos[key][0] = idx

                idx2 = idx + len(key)
                self.pos[key][1] = idx2

        for key in keys:
            if header.find(key) != -1:
                if key == 'ID':
                    if header.find('EXT') != -1:
                        self.pos['ID'][1] = header.find('EXT')

                if key == 'EXT':
                    if header.find('RESOLUTION') != -1:
                        self.pos['EXT'][1] = header.find('RESOLUTION')

                if key == 'RESOLUTION':
                    if header.find('FPS') != -1:
                        self.pos['RESOLUTION'][1] = header.find('FPS') - 1
                    elif header.find('CH') != -1:
                        self.pos['RESOLUTION'][1] = header.find('CH') - 1
                    elif header.find('|') != -1:
                        self.pos['RESOLUTION'][1] = header.find('|') - 1

                if key == 'FILESIZE':
                    if header.find('|') != -1:
                        self.pos['FILESIZE'][0] = header.find('|') + 2

                if key == 'TBR':
                    if self.pos['FILESIZE'][1] != -1:
                        self.pos['TBR'][0] = self.pos['FILESIZE'][1] + 1

                if key == 'PROTO':
                    if header.find('|', self.pos['PROTO'][0] + 1) != -1:
                        self.pos['PROTO'][1] = header.find('|', self.pos['PROTO'][0] + 1) - 1

                p = re.compile('\d+k')
                if key == 'VCODEC':
                    if header.find('VBR') != -1:
                        min_pos = 1000
                        for i in range(1, len(self.lines)):
                            s = self.lines[i][self.pos['VCODEC'][0]:]
                            m = p.search(s)
                            if m:
                                if m.start() < min_pos:
                                    min_pos = m.start()

                        if min_pos != 1000:
                            self.pos['VCODEC'][1] = self.pos['VCODEC'][0] + min_pos - 1

                    elif header.find('ACODEC') != -1:
                        self.pos['VCODEC'][1] = header.find('ACODEC') - 1

                if key == 'VBR':
                    if self.pos['VCODEC'][1] != -1:
                        self.pos['VBR'][0] = self.pos['VCODEC'][1] + 1

                if key == 'ACODEC':
                    if header.find('ABR') != -1:
                        min_pos = 1000
                        for i in range(1, len(self.lines)):
                            s = self.lines[i][self.pos['ACODEC'][0]:]
                            m = p.search(s)
                            if m:
                                if m.start() < min_pos:
                                    min_pos = m.start()

                        if min_pos != 1000:
                            self.pos['ACODEC'][1] = self.pos['ACODEC'][0] + min_pos - 1

                    elif header.find('ASR') != -1:
                        min_pos = 1000
                        for i in range(1, len(self.lines)):
                            s = self.lines[i][self.pos['ACODEC'][0]:]
                            m = p.search(s)
                            if m:
                                if m.start() < min_pos:
                                    min_pos = m.start()

                        if min_pos != 1000:
                            self.pos['ACODEC'][1] = self.pos['ACODEC'][0] + min_pos - 1

                    elif header.find('MORE INFO') != -1:
                        self.pos['ACODEC'][1] = header.find('MORE INFO') - 1

                    else:
                        self.pos['ACODEC'][1] = 1000

                if key == 'ABR':
                    if self.pos['ACODEC'][1] != -1:
                        self.pos['ABR'][0] = self.pos['ACODEC'][1] + 1

                if key == 'ASR':
                    min_pos = 1000
                    for i in range(1, len(self.lines)):
                        s = self.lines[i][self.pos['ABR'][1]:]
                        m = p.search(s)
                        if m:
                            if m.start() < min_pos:
                                min_pos = m.start()

                    if min_pos != 1000:
                        self.pos['ASR'][0] = self.pos['ABR'][1] + min_pos

    def dvlc_init(self):
        self.dvlc.SelectRow(0)
        wx.CallLater(1, self.sel_format)
        self.dvlc.SetFocus()

    def sel_format(self):
        self.onsel_format()

    def onsel_format(self, evt=None):
        if evt:
            self.cur_dvlc = 1

        if self.worker:
            if self.row_1st != -1:
                self.dvlc.SelectRow(self.row_1st)
            return
        else:
            self.btnPreview.Enable()
            self.btnUnselect.Disable()
            self.btnPreview2.Disable()
            if not self.is_live:
                self.btnDownload.Enable()
                self.cbRemux.Enable()

            self.cur_video['format'] = [None, None]
            row = self.dvlc.GetSelectedRow()
            if self.row_1st != -1:
                l = self.L3[self.row_1st]
                val = f'{l[2]} ( {l[1]}, {l[6].strip()} )' if l[6] else f'{l[2]} ( {l[1]} )'
                self.dvlc.SetValue(val, self.row_1st, 0)

            self.row_1st = row
            l = self.L3[row]
            val = f'◉ {l[2]} ( {l[1]}, {l[6].strip()} )' if l[6] else f'◉ {l[2]} ( {l[1]} )'

            self.dvlc.SetValue(val, row, 0)
            self.cur_video['format'][0 if self.L3[row][8] else 1] = self.L3[row][0]
            self.dvlc_3.DeleteAllItems()
            items = ['ID', '확장자', '해상도', '초당 프레임 수', '채널 수', '파일 크기',
                     '총 비트 전송률', '프로토콜', '비디오 코덱', '비디오 비트 전송률', '오디오 코덱',
                     '오디오 비트 전송률', '오디오 샘플 속도', '기타', '영상', '음성']

            for i in range(1, 16):
                # '프로토콜', '기타' 항목은 는 표시하지 않음
                if i in [13]:
                    continue

                val = self.L3[row][i]

                # 확장자
                if i == 1:
                    ext = self.L3[row][1]
                    if self.L3 and self.L4:
                        row = self.dvlc.GetSelectedRow()
                        row_2 = self.dvlc_2.GetSelectedRow()
                        ext_2 = self.L4[row_2][1]
                        if ext == 'm3u8' and ext_2 == 'm4a':
                            val = 'mp4'
                        elif ext == 'm3u8' and ext_2 == 'webm':
                            val = 'webm'

                    if val == 'mp4':
                        self.cbRemux.Hide()
                        self.cbRemux_visible = False
                        self.btnHelp3.Hide()
                    else:
                        if not self.is_live:
                            self.cbRemux.Show()
                            self.cbRemux_visible = True
                            self.btnHelp3.Show()

                # 초당 프레임 수
                if i == 3 and self.L3[row][3] == 'None':
                    if self.host == 'facebook':
                        val = ''
                    else:
                        pass

                # 오디오 샘플 속도
                if i == 12 and self.host not in ['ebs']:
                    val = val if val else ''

                self.dvlc_3.AppendItem([items[i], val.strip()])

            self.inner3.Show(self.bsizer3)
            self.dvlc_3.Layout()

            # 음성이 없는 경우 오디오 전용 포맷 선택
            if self.L3[row][15] == 'X':
                self.dvlc_2.Enable()
                self.display_formats2()
            else:
                self.dvlc_2.Disable()
                self.inner3.Hide(self.bsizer2)
                self.inner3.Layout()

    def onsel_format_2(self, evt=None):
        if evt:
            self.cur_dvlc = 2

        if self.worker:
            if self.row_2nd != -1:
                self.dvlc_2.SelectRow(self.row_2nd)
            return
        else:
            row = self.dvlc_2.GetSelectedRow()
            if row == -1:
                self.dvlc_2.SelectRow(self.row_2nd)
                return

            self.btnUnselect.Enable()
            self.btnPreview2.Enable()
            self.btnDownload.Enable()
            self.cbRemux.Enable()
            if self.row_2nd != -1:
                self.dvlc_2.SetValue(self.L4[self.row_2nd][1], self.row_2nd, 0)

            self.row_2nd = row
            self.dvlc_2.SetValue('◉ ' + self.L4[row][1], row, 0)
            self.cur_video['format'][1] = self.L4[row][0]
            # self.L3: 0=>ID, 1=>EXT, 2=>RESOLUTION, 3=>FPS, 4=>CH, 5=>FILESIZE, 6=>TBR, 7=>PROTO, 8=>VCODEC, 9=>VBR, 10=>ACODEC, 11=>ABR, 12=>ASR, 13=>MORE INFO, 14=>VIDEO(O/X), 15=>AUDIO(O/X), 16=>url
            # self.L4: 0=>id, 1=>ext, 2=>ch, 3=>filesize, 4=>acodec, 5=>abr, 6=>asr, 7=>url, 8=>proto
            ext = self.L3[self.row_1st][1]
            proto = self.L3[self.row_1st][7]
            ext_2 = self.L4[self.row_2nd][1]
            proto_2 = self.L4[self.row_2nd][8]
            if (ext == 'mp4' and proto == 'https') and (ext_2 == 'mp4' and proto_2 == 'm3u8'):
                ext_ = 'mp4'
            elif (ext == 'mp4' and proto == 'https') and (ext_2 == 'm4a' and proto_2 == 'https'):
                ext_ = 'mp4'
            elif (ext == 'mp4' and proto == 'm3u8') and (ext_2 == 'm4a' and proto_2 == 'https'):
                ext_ = 'mp4'
            elif (ext == 'mp4' and proto == 'm3u8') and (ext_2 == 'mp4' and proto_2 == 'm3u8'):
                ext_ = 'mp4'
            elif (ext == 'mp4' and proto == 'https') and (ext_2 == 'webm' and proto_2 == 'https'):
                ext_ = 'mkv'
            elif (ext == 'webm' and proto == 'https') and (ext_2 == 'm4a' and proto_2 == 'https'):
                ext_ = 'mkv'
            elif (ext == 'webm' and proto == 'https') and (ext_2 == 'webm' and proto_2 == 'https'):
                ext_ = 'webm'
            elif (ext == 'mp4' and proto == 'm3u8') and (ext_2 == 'webm' and proto_2 == 'https'):
                ext_ = 'webm'
            else:
                ext_ = ext

            self.ext_origin = ext_

            if ext_ == 'mp4':
                self.cbRemux.Hide()
                self.cbRemux_visible = False
                self.btnHelp3.Hide()
            else:
                if not self.is_live:
                    self.cbRemux.Show()
                    self.cbRemux_visible = True
                    self.btnHelp3.Show()

            self.dvlc_3.SetValue(ext_, 0, 1)

            # 파일 크기
            self.total_size_in_dvlc_3()

            # 채널 수
            self.dvlc_3.SetValue(self.L4[row][2].strip(), 3, 1)

            # 총 비트 전송률
            vbr = int(self.L3[self.row_1st][6].strip()[:-1])
            if self.L4[row][5].strip():
                abr = int(self.L4[row][5].strip()[:-1])
                tbr = f'{vbr + abr}k'
                self.dvlc_3.SetValue(tbr, 5, 1)

            # 오디오 코덱, 오디오 비트 전송률, 오디오 샘플 속도
            for i in range(4, 7):
                self.dvlc_3.SetValue(self.L4[row][i], i + 5, 1)

            # 음성(O/X)
            self.dvlc_3.SetValue('O', 13, 1)

    def display_formats2(self):
        self.inner3.Show(self.bsizer2)
        self.dvlc_2.Layout()
        self.sizer.Layout()
        self.dvlc_2.DeleteAllItems()

        for l in self.L4:
            drc = 'DRC' if l[0].endswith('-drc') else ''
            self.dvlc_2.AppendItem([l[1], l[5], l[3], drc])  # 1=>ext, 3=>filesize, 5=>abr

        self.cur_video['format'][1] = None
        row = self.row_2nd if self.row_2nd != -1 else 0
        self.dvlc_2.SelectRow(row)
        self.onsel_format_2()

    def onopen_dir(self, evt):
        Popen(f'explorer "{self.config["download-dir"]}"')

    def onopen_dir2(self, evt=None):
        Popen(f'explorer /select, "{self.outfile}"')

    def onopen_file(self, evt):
        os.startfile(self.outfile)

    def onhelp(self, evt):
        dlg = Help(self, 1)
        dlg.ShowModal()
        dlg.Destroy()

    def onpreview(self, evt):
        if 'format' not in self.cur_video:
            return

        row = self.dvlc.GetSelectedRow()
        url = self.L3[row][-1]
        os.startfile(url)
        self.dvlc.SetFocus()

    def onhelp2(self, evt):
        dlg = Help(self, 2)
        dlg.ShowModal()
        dlg.Destroy()

    def onhelp_it3(self, evt):
        dlg = Help(self, 3)
        dlg.ShowModal()
        dlg.Destroy()

    def onunselect_it(self, evt):
        self.cur_video['format'][1] = None
        row = self.dvlc_2.GetSelectedRow()
        self.dvlc_2.UnselectAll()
        self.dvlc_2.SetValue(self.L4[row][1], row, 0)
        row2 = self.dvlc.GetSelectedRow()
        self.dvlc_3.SetValue(self.L3[row2][1], 0, 1)  # ext
        video_size = self.L3[self.row_1st][5].replace('~', '').strip()[:-3] + 'MiB'
        self.dvlc_3.SetValue(video_size, 4, 1)
        self.dvlc_3.SetValue(self.L3[row2][6].strip()[:-1] + 'k', 5, 1)  # tbr
        for i in range(9, 12):
            self.dvlc_3.SetValue('', i, 1)

        self.dvlc_3.SetValue('X', 13, 1)
        self.btnUnselect.Disable()
        self.btnPreview2.Disable()
        self.dvlc.SetFocus()

    def onpreview2(self, evt):
        if 'format' not in self.cur_video:
            return

        pids = self.enum_procs()
        for pid in pids:
            self.enum_proc_wnds(pid)

        row = self.dvlc_2.GetSelectedRow()
        url = self.L4[row][7]
        os.startfile(url)
        self.dvlc_2.SetFocus()

    def enum_windows_proc(self, hwnd, lparam):
        if (lparam is None) or ((lparam is not None) and (win32process.GetWindowThreadProcessId(hwnd)[1] == lparam)):
            text = win32gui.GetWindowText(hwnd)
            if text:
                wstyle = win32api.GetWindowLong(hwnd, win32con.GWL_STYLE)
                if wstyle & win32con.WS_VISIBLE:
                    if re.search(self.wtxt, text):
                        win32gui.SetForegroundWindow(hwnd)
                        if not self.wsh:
                            self.wsh = win32com.client.Dispatch("WScript.Shell")

                        self.wsh.SendKeys('^w')

    def enum_proc_wnds(self, pid=None):
        win32gui.EnumWindows(self.enum_windows_proc, pid)

    @staticmethod
    def enum_procs():
        pids = win32process.EnumProcesses()
        return pids

    def oncheck_remux(self, evt):
        self.config["remuxing"] = self.cbRemux.GetValue()

    def ondownload(self, evt=None):
        if is_running(f'{YT_DLP}'):
            message = f'{TITLE}(중복 실행)에서 다운로드 진행 중입니다.\n\n다운로드를 중지하거나 다운로드 완료 후 다시 시도해주세요.'
            wx.MessageBox(f'{message}', TITLE, style=wx.ICON_WARNING)
            return

        self.status.SetLabel('')
        self.task = 'download'
        row = self.dvlc.GetSelectedRow()
        self.last_sel = (row, -1)

        # self.L3: 0=>ID, 1=>EXT, 2=>RESOLUTION, 3=>FPS, 4=>CH, 5=>FILESIZE, 6=>TBR, 7=>PROTO, 8=>VCODEC, 9=>VBR, 10=>ACODEC, 11=>ABR, 12=>ASR, 13=>MORE INFO, 14=>VIDEO(O/X), 15=>AUDIO(O/X), 16=>url
        resolution = self.L3[row][2]
        if resolution == 'audio only' and self.dvlc_2.GetSelectedRow() == -1:
            message = f'오디오 전용 포맷입니다. 오디오를 추가해주세요.\n\n '
            wx.MessageBox(message, TITLE, wx.ICON_INFORMATION | wx.OK)
            return

        tbr = self.L3[row][6].strip()
        vbr = self.L3[row][9].strip()
        abr = self.L3[row][11].strip()
        drc = ''
        # self.L4: 0=>id, 1=>ext, 2=>ch, 3=>filesize, 4=>acodec, 5=>abr, 6=>asr, 7=>url, 8=>proto
        if self.dvlc_2.GetSelectedRow() != -1:
            drc = self.L4[self.dvlc_2.GetSelectedRow()][0]

        abr_2 = ''
        if self.dvlc_2.IsEnabled() and self.dvlc_2.GetSelectedRow() != -1:
            row2 = self.dvlc_2.GetSelectedRow()
            abr_2 = self.L4[row2][5].strip()
            self.last_sel = (row, row2)

        ext = self.dvlc_3.GetValue(0, 1)

        vbr_ = f', vbr{vbr}'
        abr_ = f', abr{abr_2}' if abr_2 else f', abr{abr}'
        drc_ = ' drc' if drc.endswith('-drc') else ''
        if self.dvlc_2.IsEnabled():
            filebase = f'{self.cur_video["title"]} [{self.cur_video["id"]}] ({resolution}{vbr_}{abr_}{drc_})'
        else:
            if vbr and abr:
                filebase = f'{self.cur_video["title"]} [{self.cur_video["id"]}] ({resolution}{vbr_}{abr_}{drc_})'
            elif tbr:
                filebase = f'{self.cur_video["title"]} [{self.cur_video["id"]}] ({resolution}, tbr{tbr})'
            else:
                filebase = f'{self.cur_video["title"]} [{self.cur_video["id"]}] ({resolution})'

        filename = f'{filebase}.{ext}'

        # Windows 파일명에 금지됝 문자들 제거: \ / : * ? " < > |
        filename_ = re.sub('"', '＂', filename)
        filename_ = re.sub('\?', '？', filename_)
        filename_ = re.sub('\|', '│', filename_)
        filename_ = re.sub('/', '⧸', filename_)
        filename_ = re.sub(r'[\\:*<>]', '', filename_)
        # 한글(HWP)에서 사용하는 따옴표(“”, ‘’)는 ＂, '로 바꿈:
        filename_ = re.sub('[“”]', '＂', filename_)
        filename_ = re.sub("[‘’]", "'", filename_)
        # 띄어쓰기 연속 두 번 이상의 경우엔 한 번으로 바꿈
        filename_ = re.sub('\s{2,}', ' ', filename_)

        self.outfile = f'{self.config["download-dir"]}\\{filename_}'
        if os.path.isfile(self.outfile):
            file = os.path.split(self.outfile)[1]
            self.status.SetLabel(f'[다운로드: 같은 이름의 파일] {file}')
            self.btnOpenDir.Show()
            self.btnOpen.Show()
            self.inner2.Layout()
            message = f'[다운로드] 같은 이름의 파일이 있습니다. 덮어쓰기를 할까요?\n\n{self.outfile}'
            with wx.MessageDialog(self, message, TITLE,
                                  style=wx.YES_NO | wx.NO_DEFAULT | wx.ICON_WARNING) as messageDialog:
                if messageDialog.ShowModal() == wx.ID_YES:
                    os.remove(self.outfile)
                else:
                    return

        self.cur_video['ext'] = ext if self.task == 'download' else 'mp4'
        self.no_ffmpeg = False

        if self.L3[row][15] == 'X':
            self.gauge.SetRange(217 if self.cur_video['ext'] == 'mkv' else 213)
        else:
            self.gauge.SetRange(117 if self.cur_video['ext'] == 'mkv' else 113)

        self.btnAbort.SetFocus()
        self.set_controls(self.task)
        self.worker = WorkerThread(self)
        self.worker.daemon = True
        self.worker.start()

    def ondownload_dir(self, evt=None):
        dlg = wx.DirDialog(self, "다운로드 파일을 저장할 폴더를 선택하세요.", style=wx.DD_DIR_MUST_EXIST)
        val = dlg.ShowModal()
        path = dlg.GetPath()
        dlg.Destroy()
        if val == wx.ID_OK:
            self.config["download-dir"] = path

    def save_controls(self, arg=None):
        self.control_state['inner2.IsShown'] = self.sizer.IsShown(self.inner2)
        self.control_state['btnOpenDir.IsShown'] = self.btnOpenDir.IsShown()
        self.control_state['btnOpen.IsShown'] = self.btnOpen.IsShown()
        self.control_state['btnAbort.IsEnabled'] = self.btnAbort.IsEnabled()
        self.control_state['btnExtract.IsEnabled'] = self.btnExtract.IsEnabled()
        self.control_state['btnDownload.IsEnabled'] = self.btnDownload.IsEnabled()
        self.control_state['cbRemux.IsEnabled'] = self.cbRemux.IsEnabled()

        if arg == 'extract':
            self.control_state['inner3.IsShown'] = self.sizer.IsShown(self.inner3)
            self.control_state['inner4.IsShown'] = self.sizer.IsShown(self.inner4)
            self.control_state['btnPreview.IsEnabled'] = self.btnPreview.IsEnabled()
            self.control_state['btnPreview2.IsEnabled'] = self.btnPreview2.IsEnabled()

        elif arg in ['download', 'remux']:
            self.control_state['dvlc.IsEnabled'] = self.dvlc.IsEnabled()
            self.control_state['dvlc_2.IsEnabled'] = self.dvlc_2.IsEnabled()
            self.control_state['btnUnselect.IsEnabled'] = self.btnUnselect.IsEnabled()

    def set_controls(self, arg=None):
        if arg == 'checkversion':
            return

        if arg == 'extract':
            self.btnPreview.Disable()
            self.btnPreview2.Disable()
            self.btnOpenDir.Hide()
            self.btnOpen.Hide()
            self.sizer.Hide(self.inner3)
            self.sizer.Hide(self.inner4)

        elif arg in ['download', 'remux']:
            self.dvlc.Disable()
            self.dvlc_2.Disable()
            self.btnUnselect.Disable()

        elif arg in ['ytdlp', 'kdownloader']:
            self.menuBar.Enable(101, False)
            self.menuBar.Enable(111, False)
            self.menuBar.Enable(203, False)
            self.menuBar.Enable(205, False)

        self.save_controls(arg)
        self.sizer.Show(self.inner2)
        self.sizer.Layout()
        self.btnOpenDir.Hide()
        self.btnOpen.Hide()
        self.btnAbort.Enable()
        self.btnExtract.Disable()
        self.btnDownload.Disable()
        self.cbRemux.Disable()

    def restore_controls(self, arg=None):
        if arg == 'checkversion':
            return

        self.sizer.Show(self.inner2, self.control_state['inner2.IsShown'])
        self.btnOpenDir.Show(self.control_state['btnOpenDir.IsShown'])
        self.btnOpen.Show(self.control_state['btnOpen.IsShown'])
        self.btnAbort.Enable(self.control_state['btnAbort.IsEnabled'])
        self.btnExtract.Enable(self.control_state['btnExtract.IsEnabled'])
        self.btnDownload.Enable(self.control_state['btnDownload.IsEnabled'])
        self.cbRemux.Enable(self.control_state['cbRemux.IsEnabled'])

        if arg == 'extract':
            self.sizer.Show(self.inner3, self.control_state['inner3.IsShown'])
            self.sizer.Show(self.inner4, self.control_state['inner4.IsShown'])
            self.btnPreview.Enable(self.control_state['btnPreview.IsEnabled'])
            self.btnPreview2.Enable(self.control_state['btnPreview2.IsEnabled'])
            self.btnDownload.Disable()
            self.txtURL.SetFocus()

        elif arg in ['download', 'remux']:
            self.dvlc.Enable()
            self.dvlc_2.Enable()
            self.btnUnselect.Enable()
            self.btnOpenDir.Hide()
            self.btnOpen.Hide()
            if self.cur_dvlc == 1:
                self.dvlc.SetFocus()
            elif self.cur_dvlc == 2:
                self.dvlc_2.SetFocus()

        elif arg in ['ytdlp', 'kdownloader']:
            self.menuBar.Enable(101, True)
            self.menuBar.Enable(111, True)
            self.menuBar.Enable(203, True)
            self.menuBar.Enable(205, True)

    def onremux(self, evt):
        wildcard = '동영상 (*.mkv;*.webm;*.m3u8;*.mov;*.avi;*.wmv)|' \
                   '*.mov;*.mkv;*.webm;*.m3u8;*.avi;*.wmv|모든 파일 (*.*)|*.*'
        dlg = wx.FileDialog(self, message='파일을 선택하세요.', wildcard=wildcard,
                            style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
        val = dlg.ShowModal()
        path = dlg.GetPath()
        dlg.Destroy()
        if val == wx.ID_OK:
            self.remux_it(path)

    def remux_it(self, path):
        self.infile = path
        self.task = 'remux'

        basename = os.path.basename(self.infile)
        name = os.path.splitext(basename)[0]
        self.outfile = f'{self.config["download-dir"]}\\{name}.mp4'
        if os.path.isfile(self.outfile):
            os.remove(self.outfile)

        self.duration = ''
        self.set_controls(self.task)
        self.worker2 = WorkerThread2(self)
        self.worker2.daemon = True
        self.worker2.start()

    def onupdate_ytdlp(self, evt):
        if self.worker3:
            message = f'yt-dlp 업데이트 중입니다.\n\n' \
                      f'현재 버전: {self.ytdlp_current_version}\n\n최신 버전: {self.ytdlp_latest_version}'
            wx.MessageBox(message, TITLE, style=wx.ICON_WARNING)
            return

        if self.ytdlp_current_version == self.ytdlp_latest_version:
            message = f'yt-dlp 최신 버전 사용 중입니다.\n\n최신 버전: {self.ytdlp_latest_version}'
            wx.MessageBox(message, TITLE)
            return
        else:
            self.update_notify_ytdlp = True
            message = f'yt-dlp 최신 버전이 있습니다. 업데이트할까요?\n\n' \
                      f'현재 버전: {self.ytdlp_current_version}\n\n최신 버전: {self.ytdlp_latest_version}'

            with wx.MessageDialog(self, message, TITLE,
                                  style=wx.YES_NO | wx.ICON_INFORMATION) as messageDialog:
                if messageDialog.ShowModal() == wx.ID_YES:
                    if self.worker or self.worker2:
                        task = ''
                        if self.task == 'extract':
                            task = '분석이'
                        elif self.task == 'download':
                            task = '다운로드가'
                        elif self.task == 'remux':
                            task = '리먹싱이'

                        message = f'{task} 끝난 후에 업데이트를 진행해주세요.\n\n '
                        wx.MessageBox(message, TITLE, style=wx.ICON_WARNING)
                        return

                    self.task = 'ytdlp'
                    self.worker3 = WorkerThread3(self)
                    self.worker3.daemon = True
                    self.worker3.start()

    def onupdate_kdownloader(self, evt=None):
        if self.worker5:
            message = f'{TITLE} 업데이트 중입니다.\n\n' \
                      f'현재 버전: {VERSION}\n\n최신 버전: {self.kdownloader_latest_version}'
            wx.MessageBox(message, TITLE, style=wx.ICON_WARNING)
            return

        if VERSION == self.kdownloader_latest_version:
            message = f'{TITLE} 최신 버전 사용 중입니다.\n\n최신 버전: {self.kdownloader_latest_version}'
            wx.MessageBox(message, TITLE)
            return
        else:
            self.update_notify_kdownloader = True
            message = f'{TITLE} 최신 버전이 있습니다. 업데이트할까요?\n\n' \
                      f'현재 버전: {VERSION}\n\n최신 버전: {self.kdownloader_latest_version}'

            with wx.MessageDialog(self, message, TITLE,
                                  style=wx.YES_NO | wx.ICON_INFORMATION) as messageDialog:
                if messageDialog.ShowModal() == wx.ID_YES:
                    if self.worker or self.worker2:
                        task = ''
                        if self.task == 'extract':
                            task = '분석이'
                        elif self.task == 'download':
                            task = '다운로드가'
                        elif self.task == 'remux':
                            task = '리먹싱이'

                        message = f'{task} 끝난 후에 업데이트를 진행해주세요.\n\n '
                        wx.MessageBox(message, TITLE, style=wx.ICON_WARNING)
                        return

                    self.task = 'kdownloader'
                    self.worker5 = WorkerThread3(self)
                    self.worker5.daemon = True
                    self.worker5.start()

    @staticmethod
    def en2ko(s):
        s = s.replace('Requesting header', '헤더 요청') \
            .replace('Downloading Page', '페이지 다운로드') \
            .replace('Downloading webpage', '웹페이지 다운로드') \
            .replace('Downloading player', '플레이어 다운로드') \
            .replace('Downloading m3u8 information', 'm3u8 정보 다운로드') \
            .replace('Downloading m3u8 manifest', 'm3u8 매니페스트 다운로드') \
            .replace('Downloading MPD manifest', 'MPD 매니페스트 다운로드') \
            .replace('Downloading jwt token', 'jwt 토큰 다운로드') \
            .replace('Downloading akfire_interconnect_quic m3u8 information', 'akfire_interconnect_quic m3u8 정보 다운로드') \
            .replace('Downloading fastly_skyfire m3u8 information', 'fastly_skyfire m3u8 정보 다운로드') \
            .replace('Downloading google_mediacdn m3u8 information', 'google_mediacdn m3u8 정보 다운로드') \
            .replace('Downloading akfire_interconnect_quic MPD information', 'akfire_interconnect_quic MPD 정보 다운로드') \
            .replace('Downloading fastly_skyfire MPD information', 'fastly_skyfire MPD 정보 다운로드') \
            .replace('Downloading google_mediacdn MPD information', 'google_mediacdn MPD 정보 다운로드') \
            .replace('Downloading part 1 m3u8 information', 'part 1 m3u8 정보 다운로드') \
            .replace('Extracting information', '정보 추출') \
            .replace('Extracting URL', 'URL 추출') \
            .replace('Downloading API JSON', 'API JSON 다운로드') \
            .replace('Downloading android player API JSON', '안드로이드 플레이어 API JSON 다운로드') \
            .replace('Downloading ios player API JSON', 'iOS 플레이어 API JSON 다운로드') \
            .replace('Downloading tv player API JSON', 'tv 플레이어 API JSON 다운로드') \
            .replace('Downloading web creator player API JSON', '웹 크리에이터 플레이어 API JSON 다운로드') \
            .replace('Downloading JSON metadata', 'JSON 메타데이터 다운로드') \
            .replace('Downloading 1 format(s)', '포맷 다운로드') \
            .replace('Downloading video info', '비디오 정보 다운로드') \
            .replace('Downloading video URL', '비디오 URL 다운로드') \
            .replace('Deleting existing file', '다운로드할 파일과 같은 이름의 기존 파일 삭제') \
            .replace('Deleting original file', '다운로드한 원파일 삭제') \
            .replace('Following redirect to', '다음으로 리다이렉션') \
            .replace('Total fragments', '전체 조각 수') \
            .replace('fragment not', '조각을 찾을 수 없습니다.')

        return s

    def total_size_in_dvlc_3(self):
        video_unit = self.L3[self.row_1st][5].replace('~', '').strip()[-3:]
        video_size = float(self.L3[self.row_1st][5].replace('~', '').strip()[:-3])
        video_size_kib = 0
        if video_unit == 'GiB':
            video_size_kib = video_size * 1048576
        elif video_unit == 'MiB':
            video_size_kib = video_size * 1024
        elif video_unit == 'KiB':
            video_size_kib = video_size

        if self.L3[self.row_1st][15] == 'X':
            flag = '~' if ('~' in self.L3[self.row_1st][5] or '~' in self.L4[self.row_2nd][3]) else ''
            audio_unit = self.L4[self.row_2nd][3].replace('~', '').strip()[-3:]
            audio_size = float(self.L4[self.row_2nd][3].replace('~', '').strip()[:-3])
            audio_size_kib = 0
            if audio_unit == 'GiB':
                audio_size_kib = audio_size * 1048576
            elif audio_unit == 'MiB':
                audio_size_kib = audio_size * 1024
            elif audio_unit == 'KiB':
                audio_size_kib = audio_size

            total_size_kib = video_size_kib + audio_size_kib

        else:
            flag = '~' if '~' in self.L3[self.row_1st][5] else ''
            total_size_kib = video_size_kib

        if total_size_kib >= 1048576:
            unit = 'GiB'
            total_size = f'{round(total_size_kib / 1048576, 2):.2f}'

        elif total_size_kib >= 1024:
            unit = 'MiB'
            total_size = f'{round(total_size_kib / 1024, 2):.2f}'

        else:
            unit = 'KiB'
            total_size = f'{round(total_size_kib, 2):.2f}'

        self.dvlc_3.SetValue(f'{flag}{total_size}{unit}', 4, 1)

    def onabout(self, evt):
        dlg = Help(self, 9)
        dlg.ShowModal()

    def onclose(self, evt):
        self.Close()

    def onwindow_close(self, evt):
        try:
            with open('config.pickle', 'wb') as f:
                pickle.dump(self.config, f)
        except Exception as e:
            print(e)

        self.cleanup()

        try:
            self.Destroy()
        except Exception as e:
            print(e)

    def cleanup(self):
        for file in self.tempfiles:
            if os.path.isfile(file):
                try:
                    os.remove(file)
                except Exception as e:
                    print(e)

        if self.proc:
            Popen(f'TASKKILL /F /PID {self.proc.pid} /T'.split(), creationflags=0x08000000)

        procs = [proc for proc in psutil.process_iter(['name', 'pid'])
                 if proc.info['name'] in ['explorer.exe', YT_DLP]]
        for proc in procs:
            if  proc.info['pid'] not in self.pids_explorer_existing:
                try:
                    proc.terminate()
                except Exception as e:
                    print(e)


if __name__ == '__main__':
    app = wx.App()
    frame = VideoDownloader()
    frame.Show()
    app.MainLoop()
