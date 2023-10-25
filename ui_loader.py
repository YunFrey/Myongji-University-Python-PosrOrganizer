
#PyQt와 달리 PySide6에서는 uic를 통해 클래스에 ui파일을 곧바로 로드할 수 없습니다.
# PySide에서는 아래 코드를 통해 클래스에 ui를 할당할 수 없음
# QtUiTools.QUiLoader().load('MainMenu.ui', self)
# 그래서
# https://robonobodojo.wordpress.com/2017/10/03/loading-a-pyside-ui-via-a-class/ 를 참조하여 pyside 의 ui_loader를 재정의합니다.

from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QMetaObject

class UiLoader(QUiLoader):
    def __init__(self, base_instance):
        QUiLoader.__init__(self, base_instance)
        self.base_instance = base_instance

    def createWidget(self, class_name, parent=None, name=''):
        if parent is None and self.base_instance:
            return self.base_instance
        else:
            # create a new widget for child widgets
            widget = QUiLoader.createWidget(self, class_name, parent, name)
            if self.base_instance:
                setattr(self.base_instance, name, widget)
            return widget

def load_ui(ui_file, base_instance=None):
    loader = UiLoader(base_instance)
    widget = loader.load(ui_file)
    QMetaObject.connectSlotsByName(widget)
    return widget