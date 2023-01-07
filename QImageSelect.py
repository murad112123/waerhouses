from typing import Union
from PyQt5.QtCore import QRect, QSize, Qt
from PyQt5.QtGui import QPixmap, QTransform
from PyQt5.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QRubberBand, QVBoxLayout, QWidget


class QImageSelect(QDialog):
    """Image Cropper Dialog.

    Parameters
    ----------
    title : str
        Dialog title
    image : Union[str, QPixmap]
        Image path or qpixmap
    maximum_size : QSize
        Maximum dialog size
    parent : QWidget, optional
        Parent widget, by default None

    Returns
    -------
    QPixmap
        Cropped qpixmap
    """

    result = None
    ROTATE_BTN_MAX_WIDTH = 35

    @classmethod
    def spawn(cls, title: str, image: Union[str, QPixmap], maximum_size: QSize, parent: QWidget = None) -> QPixmap:
        """Create instance of QImageSelect, execute it and return cropped qpixmap

        Parameters
        ----------
        title : str
            Dialog title
        image : Union[str, QPixmap]
            Image path or qpixmap
        maximum_size : QSize
            Maximum dialog size
        parent : QWidget, optional
            Parent widget, by default None

        Returns
        -------
        QPixmap
            Cropped qpixmap
        """
        dialog = cls(title=title, image=image, maximum_size=maximum_size, parent=parent)
        if not dialog.exec():
            return None
        rect = dialog.pic.selected_rect
        if not rect:
            rect = dialog.pixmap.rect()
        return dialog.pixmap.copy(rect.x(), rect.y(), rect.width(), rect.height())

    def __init__(self, title: str, image: Union[str, QPixmap], maximum_size: QSize, parent: QWidget = None):
        """Create instance of QImageSelect."""

        super(QImageSelect, self).__init__(parent)
        self.initUI(title, image, maximum_size)

    def initUI(self, title: str, image: Union[str, QPixmap], maximum_size: QSize):
        """Initialize UI."""

        self.setWindowTitle(title)

        self.bg_layout = QVBoxLayout(self)
        self.bg_layout.setContentsMargins(0, 0, 0, 0)
        self.pic = self.QImageLabel(self)
        self.btn_container = QWidget(self)
        self.btn_layout = QHBoxLayout(self.btn_container)
        self.btn_layout.setContentsMargins(0, 0, 0, 0)
        self.btn_ok = QPushButton("OK", self.btn_container)
        self.btn_rotate = QPushButton("↻", self.btn_container)
        self.btn_rotate.setMaximumWidth(self.ROTATE_BTN_MAX_WIDTH)
        self.btn_layout.addWidget(self.btn_rotate)
        self.btn_layout.addWidget(self.btn_ok)
        self.bg_layout.addWidget(self.btn_container)
        self.bg_layout.addWidget(self.pic)

        self.pixmap = QPixmap(image)
        self.pixmap = self.pixmap.scaled(maximum_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.adjust_image(0)

        self.setMouseTracking(True)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_rotate.clicked.connect(lambda: self.adjust_image())

    def adjust_image(self, rotate_degree: int = 90):
        """Rotate image and resize dialog."""

        transform = QTransform().rotate(rotate_degree)
        self.pixmap = self.pixmap.transformed(transform)
        width = self.pixmap.width()
        hight = self.pixmap.height()
        self.setFixedSize(width, hight + 45)
        self.pic.setPixmap(self.pixmap)

    class QImageLabel(QLabel):
        """Customized QLabel with rubberband support.

        Parameters
        ----------
        parent : QWidget
            Parent widget
        """

        def __init__(self, parent: QWidget):
            """Create instance of QImageLabel."""

            super().__init__(parent=parent)
            self.rubberband = QRubberBand(QRubberBand.Rectangle, self)
            self.selected_rect = QRect()

        def mousePressEvent(self, event):
            if event.button() == Qt.LeftButton:
                self.origin = event.pos()
                self.rubberband.setGeometry(QRect(self.origin, QSize()))
                self.rubberband.show()

        def mouseMoveEvent(self, event):
            if self.rubberband.isVisible():
                self.rubberband.setGeometry(QRect(self.origin, event.pos()).normalized() & self.rect())

        def mouseReleaseEvent(self, event):
            if self.rubberband.isVisible() and event.button() == Qt.LeftButton:
                self.selected_rect = self.rubberband.geometry().normalized()
