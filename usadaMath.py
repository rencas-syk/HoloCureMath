from mss import mss
import win32gui
import cv2
import numpy as np
import os
import sys
from PyQt6.QtWidgets import QLabel, QMainWindow, QWidget, QApplication
from PyQt6.QtGui import QFont, QFontDatabase
from PyQt6.QtCore import Qt
from threading import Thread
from time import sleep
from pyautogui import locate




def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception as e:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)



class Runner():
    def __init__(self):
        self.display = '...'
        self.displayMessageSubscribers = set()
        pass

    def subscribe(self, obj):
        self.displayMessageSubscribers.add(obj)
    
    def unsubscribe(self, obj):
        self.displayMessageSubscribers.remove(obj)

    def sendDisplayMessage(self):
        for sub in self.displayMessageSubscribers:
            sub.receiveDisplayMessage(self.display)

    
    def _preprocess(self, frame:np.ndarray) -> np.ndarray:
        grayscale = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        _, binarized = cv2.threshold(grayscale,220,255,cv2.THRESH_BINARY)
        return binarized

    def collect_templates(self, template_path: str):
        template_dir = os.scandir(resource_path(template_path))
        templates = []
        names = []
        for file in template_dir:
            file_path = resource_path(os.path.join(template_path, file.name))
            templates.append(cv2.imread(file_path, cv2.IMREAD_GRAYSCALE))
            names.append(file.name.split('.')[0])
        return templates, names

    def resize_templates(self, templates, scalar):
        if scalar < 1:
            return [cv2.resize(t,None, fx = scalar, fy = scalar, interpolation = cv2.INTER_AREA) for t in templates]
        elif scalar > 1:
            return [cv2.resize(t,None, fx = scalar, fy = scalar, interpolation = cv2.INTER_CUBIC) for t in templates]
        else:
            return templates

    def buildFormula(self, img: np.ndarray, templates: np.ndarray, names: str, method = cv2.TM_CCOEFF_NORMED):
        formula = []
        for template, name in zip(templates, names):
            res = cv2.matchTemplate(img,template,method)
            threshold = 0.97
            loc = np.where( res >= threshold)
            prev_pt = 0
            for pt in zip(*loc[::-1]):
                if abs(pt[0] - prev_pt) < 7:
                    continue
                if 'mult' in name:
                    symbol = '*'
                elif 'div'  in name:
                    symbol = '//'
                else:
                    symbol = name
                formula.append((symbol,pt[0]))
                prev_pt = pt[0]
        formula = [y[0] for y in sorted(formula, key=lambda x: x[1])]
        return ''.join(formula)

    def updateRegion(self, hwnd):
        window = win32gui.GetClientRect(hwnd)
        width = window[2]
        height = window[3]
        rel_rect = (0.195, 0.127, 0.797, 0.183)

        abs_pos1 = (int(rel_rect[0] * width), int(rel_rect[1]*height))
        abs_pos2 = (int(rel_rect[2] * width), int(rel_rect[3]*height))
        pos1 = win32gui.ClientToScreen(hwnd, abs_pos1)
        pos2 = win32gui.ClientToScreen(hwnd, abs_pos2)
        region = tuple([*pos1, *pos2])

        return (region, width, height)

    def _start(self):
        self.t = Thread(target=self._run, args=())
        self.active = True
        self.t.start()
    
    def _run(self):
        sct = mss()
        templates, names = self.collect_templates('img')
        prev_height = 1080
        
        while self.active:
            try:
                hwnd = win32gui.FindWindowEx(None, None, None, 'HoloCure')
                region, width, height = self.updateRegion(hwnd)
                if height != prev_height:
                    templates = self.resize_templates(templates=templates, scalar= height / 1080)
                    prev_height = height
            except Exception as e:
                hwnd = None
                self.display = ('HoloCure\nnot found!')
                continue
            img = np.array(sct.grab(region))
            img = self._preprocess(img)
            formula = self.buildFormula(img, templates=templates, names=names)
            try:
                res = eval(formula)
            except Exception as e:
                res = '...'
                print(e)

            self.display = str(res)

            self.sendDisplayMessage()
        
    def _stop(self):
        self.active = False
        self.t.join()

class OverlayLabel(QLabel):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.WindowTransparentForInput | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(73,39)
        self.setFont(QFont('Joystix Monospace', int(22*self.width()/63)))
        self.setStyleSheet("color: white; \
                           background-color: rgba(127,0, 0,127);")
        self.setText('...')
        # self._start()

    

    def updateRegion(hwnd):
        window = win32gui.GetClientRect(hwnd)
        abs_pos1 = window[0], window[1]
        abs_pos2 = window[2], window[3]
        pos1 = win32gui.ClientToScreen(hwnd,abs_pos1)
        pos2 = win32gui.ClientToScreen(hwnd,abs_pos2)
        region = tuple([*pos1, *pos2])

        return region
    
    def _preprocess(self, frame:np.ndarray) -> np.ndarray:
        grayscale = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        _, binarized = cv2.threshold(grayscale,220,255,cv2.THRESH_BINARY)
        return binarized
    
    def resize_template(self, template, scalar):
        if scalar < 1:
            return cv2.resize(template,None, fx = scalar, fy = scalar, interpolation = cv2.INTER_AREA)
        elif scalar > 1:
            return cv2.resize(template,None, fx = scalar, fy = scalar, interpolation = cv2.INTER_CUBIC)
        else:
            return template

    def _start(self):
        self.t = Thread(target=self._run, args=())
        self.active = True
        self.t.start()        
    
    def _run(self):
        sct = mss()
        template = resource_path('pag.png')

        while self.active:
            try:
                hwnd = win32gui.FindWindowEx(None, None, None, 'HoloCure')
                region, width, height = self.updateRegion(hwnd)
                if height != prev_height:
                    template = self.resize_template(template=template, scalar= height / 1080)
                    prev_height = height
            except Exception as e:
                hwnd = None
                self.display = ('...')
                self.hide()
                continue
            img = np.array(sct.grab(region))
            img = self._preprocess(img)
            try:
                box = locate(template, img)
                abs_pos1 = win32gui.ClientToScreen(hwnd, (box.left, box.top))
                self.setFixedSize(box.width+10, box.height)
                self.move(abs_pos1[0], abs_pos1[1])
            except:
                self.hide()

    def _stop(self):
        self.active = False
        self.t.join()
    
    def closeEvent(self, a0):
        self._stop()
        return super().closeEvent(a0)

    def showEvent(self, a0):
        self._start()
        self.move(0,0)
        return super().showEvent(a0)



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('HoloCure MathSolver')
        self.setFixedSize(250,100)
        QFontDatabase.addApplicationFont(resource_path('joystix.monospace-regular.otf'))

        self.cwidget = QLabel()
        self.cwidget.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cwidget.setFont(QFont('Joystix Monospace', 20))

        self.content = '...'
        self.cwidget.setText(self.content)

        self.setCentralWidget(self.cwidget)

        # self.overlay = OverlayLabel()
        # self.overlay.show()

        self.show()



        self.runner = Runner()
        self.runner.subscribe(self)
        self.runner._start()


    def receiveDisplayMessage(self, displayMessage):
        self.cwidget.setText(displayMessage)
        #self.overlay.setText(displayMessage)
    


    def closeEvent(self, event):
        #self.overlay.close()
        self.runner._stop()
        super().closeEvent(event)
        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    guiwindow = MainWindow()
    sys.exit(app.exec())