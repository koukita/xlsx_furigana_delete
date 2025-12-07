# -*- coding: utf-8 -*-
from qgis.PyQt import QtWidgets
from qgis.core import QgsProject

class XlsxFuriganaDeleteDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Excel ふりがな削除")

        # レイヤ選択ラベル
        label = QtWidgets.QLabel("対象レイヤを選択：")

        # レイヤコンボボックス
        self.layer_combo = QtWidgets.QComboBox()
        layers = QgsProject.instance().mapLayers().values()
        for layer in layers:
            self.layer_combo.addItem(layer.name())

        # OK/Cancel ボタン
        self.buttonBox = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )

        # レイアウト作成
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(label)
        layout.addWidget(self.layer_combo)
        layout.addWidget(self.buttonBox)
        self.setLayout(layout)
