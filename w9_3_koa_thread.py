import logging
import pythoncom

from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

def qthreadId():
    return int(QThread.currentThreadId())
    # return QThread.currentThread()

def qthread_desc():
    return f'{qthreadId()}/{QThread.currentThread().objectName()}'

class MainThreadWidget(QLabel):
    def __init__(self):
        super().__init__()
        self.resize(400,400)

        logging.info(f'{qthreadId()}: MainThreadWidget')

        self.timer = QTimer(self)
        self.timer.start(5*1000)
        self.timer.timeout.connect(lambda: logging.info(f'{qthread_desc()}: MainThreadWidget') )

class Worker(QObject):
    def __init__(self):
        super().__init__()
        self.ocx: QAxWidget= None
        logging.info(f'{qthread_desc()} Worker')

    def __del__(self):
        logging.info(f'~Worker')

    def on_started(self):
        QTimer.singleShot(1000, self.post_init)
        # ...

    def post_init(self):
        logging.info(f'{qthread_desc()} 1.post_init')
        pythoncom.CoInitialize()
        logging.info(f'{qthread_desc()} 2.post_init')

        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        logging.info(f'{qthread_desc()} 3.post_init')

        self.ocx.OnEventConnect.connect(self.OnEventConnect)
        self.ocx.OnReceiveTrData.connect(self.OnReceiveTrData)

        logging.info(f'{qthread_desc()} 1.CommConnect')
        self.ocx.dynamicCall('CommConnect')
        logging.info(f'{qthread_desc()} 2.CommConnect')
        pass

    def on_finished(self):
        logging.info(f'{qthread_desc()} 1.on_finished')
        self.ocx = None

    def OnEventConnect(self, nErrCode: int) -> None:
        from functools import partial

        logging.info(f"{qthread_desc()} [OnEventConnect] err:{nErrCode}")
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.setSingleShot(False)
        self.timer.timeout.connect(partial(self.opt10001, '005930'))
        self.timer.start()
        # self.opt10001('005930')

    def OnReceiveTrData(self, sScrNo: str, sRQName: str, sTrCode: str, sRecordName: str, sPrevNext: str,
                        nDataLength: int, sErrorCode: str, sMessage: str, sSplmMsg: str) -> None:
        logging.info(f"{qthread_desc()} [OnReceiveTrData] sRQName:{sRQName}")
        if sTrCode == 'opt10001':
            r = self.ocx.dynamicCall('GetCommData(QString,QString,int,QString)', sTrCode, sRQName, 0, '종목코드')
            logging.info(f' 종목코드:{r!r}')
            r = self.ocx.dynamicCall('GetCommData(QString,QString,int,QString)', sTrCode, sRQName, 0, '종목명')
            logging.info(f' 종목명:{r!r}')
            r = self.ocx.dynamicCall('GetCommData(QString,QString,int,QString)', sTrCode, sRQName, 0, '현재가')
            logging.info(f' 현재가:{r!r}')

    @pyqtSlot(str)
    def opt10001(self, code: str):
        r = self.ocx.dynamicCall('SetInputValue(QString,QString)', '종목코드', code)
        r = self.ocx.dynamicCall('CommRqData(QString,QString,int,QString)', 'RQName', 'opt10001', 0, '0000')

    @pyqtSlot()
    def test1(self):
     logging.info(f'{qthread_desc()} 1.test1')

    @pyqtSlot(int,str)
    def test2(self, i, s):
     logging.info(f'{qthread_desc()} 1.test2 i:{i}, s:{s}')

    @pyqtSlot(object)
    def test3(self, obj):
     logging.info(f'{qthread_desc()} 1.test3 obj:{obj}')


def test_qthread():
    app = QApplication([])
    w = MainThreadWidget()
    w.show()

    # sub thread Kiwoom ocx
    th = QThread()
    sub_thread_worker = Worker()
    sub_thread_worker.moveToThread(th)
    th.started.connect(sub_thread_worker.on_started)
    th.finished.connect(sub_thread_worker.on_finished)
    th.start()

    # main thread Kiwoom ocx
    main_thread_worker = Worker()
    QTimer.singleShot(5000, main_thread_worker.on_started)

    # QMetaObject.invokeMethod(worker, "test", Qt.QueuedConnection, Q_ARG("PyQt_PyObject", QThread.currentThread()))
    QMetaObject.invokeMethod(sub_thread_worker, "test1", Qt.QueuedConnection)
    QMetaObject.invokeMethod(sub_thread_worker, "test2", Qt.QueuedConnection, Q_ARG(int, 1), Q_ARG(str, 'hello'))
    QMetaObject.invokeMethod(sub_thread_worker, "test3", Qt.QueuedConnection, Q_ARG(object, [1,2,3]))

    logging.info(f'1')
    app.exec_()

    logging.info(f'2.quit')
    th.quit()
    logging.info(f'3.wait')
    th.wait()
    # th.metaObject().invokeMethod(th, 'do_finish', Qt.QueuedConnection)
    logging.info(f'4')

if __name__ == '__main__':
    LOG_FORMAT = '%(pathname)s:%(lineno)03d | %(asctime)s | %(levelname)s | %(message)s'
    LOG_LEVEL = logging.INFO  # DEBUG(10), INFO(20), (0~50)

    logging.basicConfig(format=LOG_FORMAT, level=LOG_LEVEL)
    test_qthread()