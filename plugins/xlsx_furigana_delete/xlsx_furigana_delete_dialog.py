# -*- coding: utf-8 -*-
from qgis.PyQt import QtWidgets
from qgis.core import QgsProject

class XlsxFuriganaDeleteDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Excel ふりがな削除")

        # 注意書きラベル
        w_label = QtWidgets.QLabel("注意！ \n テーブル結合、リレーションを行っている場合は解除されます。 \n 別のプロジェクトで実行してください。")

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
        layout.addWidget(w_label)
        layout.addWidget(label)
        layout.addWidget(self.layer_combo)
        layout.addWidget(self.buttonBox)
        self.setLayout(layout)
