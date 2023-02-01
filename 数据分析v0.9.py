# 制作者：健二先生
# 邮箱：dingjianyi0508@gmail.com

import sys
import warnings
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from PyQt5.QtCore import QCoreApplication
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox, QInputDialog
from PyQt5 import QtCore, QtGui, QtWidgets
from icecream import ic
from factor_analyzer import FactorAnalyzer
from factor_analyzer.factor_analyzer import calculate_bartlett_sphericity, calculate_kmo

plt.rcParams['font.sans-serif'] = ['SimHei']
ic.disable()
warnings.filterwarnings('ignore')


def handle_title(title: str):
    if title == '' or len(title) < 1:
        return title
    elif ' ' in title or len(title) < 1:
        title = title.split(' ')[0]
        while title[-1].isalpha():
            title = title[:-1]
            if len(title) == 1:
                return title
    elif title[-1].isalpha():
        while title[-1].isalpha():
            title = title[:-1]
            if len(title) == 1:
                return title
    return title


# 计算alpha值
def calculate_alpha(df: pd.DataFrame):
    # 总分变异量计算
    total_row = df.sum(axis=1)
    sy = total_row.var()

    # 题项变异量计算
    var_column = df.var()
    si = var_column.sum()

    # 克隆巴赫alpha计算
    alpha = (len(df.columns) / (len(df.columns) - 1)) * ((sy - si) / sy)
    return alpha


# 选择出量表的col
def chose_col(col_list: list):
    chosen_col = []
    for col in col_list:
        if col[0].isalpha() and col[-1].isnumeric():
            chosen_col.append(col)
    else:
        pass
    return chosen_col


# 获取标题集合
def get_col_name(some_list: list):
    col_list = []
    for col in some_list:
        while col[-1].isnumeric():
            col = col[:-1]
        col_list.append(col)
    return set(col_list)


# 将不同的潜变量问项分组
def agg_questions(df: pd.DataFrame, mode: int):
    # 初始化结果字典
    col_dict = {}
    alpha_dict = {}

    # 提取出所有标题字符
    titles = set(map(handle_title, df.columns.tolist()))

    # 对集合排序
    titles = chose_col(sorted(list(titles)))

    col_name = sorted(get_col_name(titles))

    for item in col_name:
        questions = []
        for question in df.columns.tolist():
            if handle_title(question)[0] == item:
                questions.append(question)
        col_dict[item] = questions

    # 计算各个题项的alpha值
    for key in col_dict.keys():
        alpha = calculate_alpha(df[col_dict[key]])
        alpha_dict[key] = round(alpha, 3)

    if mode == 1:
        return col_name, col_dict
    if mode == 2:
        return alpha_dict


# 按条显示字典
def pprint_dict(to_print: dict):
    dict_str = ''
    for item in to_print:
        dict_str = dict_str + f'{item}' + ' : ' f'{to_print[item]}' + '\n'
    return dict_str


def factor_analyze(df: pd.DataFrame):
    fa = FactorAnalyzer(df.shape[1], rotation=None)
    fa.fit(df)

    ev, v = fa.get_eigenvalues()

    # 使用特征大于1的规则
    num_facter = len(ev[ev >= 1])
    fa = FactorAnalyzer(num_facter, rotation="varimax")
    fa.fit(df)
    return num_facter, fa.loadings_


class Ui_main_window(QtWidgets.QWidget):
    _data = None
    _temp_data = None
    _save_dir = None
    _col_dict = None
    _alpha_dict = None
    _factor_matrix = None
    _calculate_data = None
    _col_names = None
    _all_questions = None

    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def setupUi(self, main_window):
        main_window.setObjectName("main_window")
        main_window.resize(777, 535)
        main_window.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.tab_widget = QtWidgets.QTabWidget(main_window)
        self.tab_widget.setEnabled(True)
        self.tab_widget.setGeometry(QtCore.QRect(10, 10, 761, 491))
        self.tab_widget.setDocumentMode(False)
        self.tab_widget.setTabsClosable(False)
        self.tab_widget.setMovable(False)
        self.tab_widget.setTabBarAutoHide(False)
        self.tab_widget.setObjectName("tab_widget")
        self.data_load = QtWidgets.QWidget()
        self.data_load.setObjectName("data_load")
        self.btn_load_data = QtWidgets.QPushButton(self.data_load)
        self.btn_load_data.setGeometry(QtCore.QRect(560, 430, 75, 23))
        self.btn_load_data.setMinimumSize(QtCore.QSize(75, 23))
        self.btn_load_data.setObjectName("btn_load_data")
        self.show_current_data = QtWidgets.QTableWidget(self.data_load)
        self.show_current_data.setGeometry(QtCore.QRect(20, 20, 721, 401))
        self.show_current_data.setObjectName("show_current_data")
        self.show_current_data.setColumnCount(0)
        self.show_current_data.setRowCount(0)
        self.btn_save_data = QtWidgets.QPushButton(self.data_load)
        self.btn_save_data.setGeometry(QtCore.QRect(650, 430, 75, 23))
        self.btn_save_data.setObjectName("btn_save_data")
        self.btn_reset_data = QtWidgets.QPushButton(self.data_load)
        self.btn_reset_data.setGeometry(QtCore.QRect(40, 430, 75, 23))
        self.btn_reset_data.setObjectName("btn_reset_data")
        self.tab_widget.addTab(self.data_load, "")
        self.pre_test = QtWidgets.QWidget()
        self.pre_test.setObjectName("pre_test")
        self.tag_2_show_data = QtWidgets.QTextBrowser(self.pre_test)
        self.tag_2_show_data.setGeometry(QtCore.QRect(20, 40, 281, 411))
        self.tag_2_show_data.setObjectName("tag_2_show_data")
        self.lbl_tag_2_titile_1 = QtWidgets.QLabel(self.pre_test)
        self.lbl_tag_2_titile_1.setGeometry(QtCore.QRect(120, 10, 51, 21))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(10)
        self.lbl_tag_2_titile_1.setFont(font)
        self.lbl_tag_2_titile_1.setObjectName("lbl_tag_2_titile_1")
        self.line = QtWidgets.QFrame(self.pre_test)
        self.line.setGeometry(QtCore.QRect(320, 130, 411, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.lbl_tag_2_title_2 = QtWidgets.QLabel(self.pre_test)
        self.lbl_tag_2_title_2.setGeometry(QtCore.QRect(480, 20, 71, 20))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(9)
        self.lbl_tag_2_title_2.setFont(font)
        self.lbl_tag_2_title_2.setObjectName("lbl_tag_2_title_2")
        self.btn_cut_data = QtWidgets.QPushButton(self.pre_test)
        self.btn_cut_data.setGeometry(QtCore.QRect(450, 70, 141, 21))
        self.btn_cut_data.setObjectName("btn_cut_data")
        self.label_alpha = QtWidgets.QLabel(self.pre_test)
        self.label_alpha.setGeometry(QtCore.QRect(390, 150, 81, 20))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(9)
        self.label_alpha.setFont(font)
        self.label_alpha.setObjectName("label_alpha")
        self.btn_count_alpha = QtWidgets.QPushButton(self.pre_test)
        self.btn_count_alpha.setGeometry(QtCore.QRect(360, 170, 131, 31))
        self.btn_count_alpha.setObjectName("btn_count_alpha")
        self.label_factor = QtWidgets.QLabel(self.pre_test)
        self.label_factor.setGeometry(QtCore.QRect(380, 230, 91, 21))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(9)
        self.label_factor.setFont(font)
        self.label_factor.setObjectName("label_factor")
        self.btn_count_factor = QtWidgets.QPushButton(self.pre_test)
        self.btn_count_factor.setGeometry(QtCore.QRect(360, 250, 131, 31))
        self.btn_count_factor.setObjectName("btn_count_factor")
        self.label_check_value = QtWidgets.QLabel(self.pre_test)
        self.label_check_value.setGeometry(QtCore.QRect(600, 320, 71, 20))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(9)
        self.label_check_value.setFont(font)
        self.label_check_value.setObjectName("label_check_value")
        self.btn_check_value = QtWidgets.QPushButton(self.pre_test)
        self.btn_check_value.setGeometry(QtCore.QRect(570, 340, 131, 31))
        self.btn_check_value.setObjectName("btn_check_value")
        self.line_2 = QtWidgets.QFrame(self.pre_test)
        self.line_2.setGeometry(QtCore.QRect(320, 210, 411, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.ckb_ignore_small = QtWidgets.QCheckBox(self.pre_test)
        self.ckb_ignore_small.setGeometry(QtCore.QRect(540, 230, 161, 21))
        font = QtGui.QFont()
        font.setFamily("黑体")
        self.ckb_ignore_small.setFont(font)
        self.ckb_ignore_small.setObjectName("ckb_ignore_small")
        self.ety_small_limit = QtWidgets.QLineEdit(self.pre_test)
        self.ety_small_limit.setGeometry(QtCore.QRect(580, 260, 113, 20))
        self.ety_small_limit.setObjectName("ety_small_limit")
        self.label_small_limit = QtWidgets.QLabel(self.pre_test)
        self.label_small_limit.setGeometry(QtCore.QRect(530, 260, 54, 21))
        font = QtGui.QFont()
        font.setFamily("黑体")
        self.label_small_limit.setFont(font)
        self.label_small_limit.setObjectName("label_small_limit")
        self.line_3 = QtWidgets.QFrame(self.pre_test)
        self.line_3.setGeometry(QtCore.QRect(320, 290, 411, 20))
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.ety_col_check = QtWidgets.QLineEdit(self.pre_test)
        self.ety_col_check.setGeometry(QtCore.QRect(360, 340, 131, 31))
        self.ety_col_check.setObjectName("ety_col_check")
        self.label_col_tocheck = QtWidgets.QLabel(self.pre_test)
        self.label_col_tocheck.setGeometry(QtCore.QRect(340, 310, 211, 21))
        font = QtGui.QFont()
        font.setFamily("黑体")
        self.label_col_tocheck.setFont(font)
        self.label_col_tocheck.setObjectName("label_col_tocheck")
        self.tab_widget.addTab(self.pre_test, "")
        self.create_data = QtWidgets.QWidget()
        self.create_data.setObjectName("create_data")
        self.tab_widget.addTab(self.create_data, "")
        self.lbl_data_show = QtWidgets.QLabel(main_window)
        self.lbl_data_show.setGeometry(QtCore.QRect(520, 510, 251, 21))
        self.lbl_data_show.setObjectName("lbl_data_show")
        self.label = QtWidgets.QLabel(main_window)
        self.label.setGeometry(QtCore.QRect(20, 510, 54, 12))
        self.label.setObjectName("label")

        self.retranslateUi(main_window)
        self.tab_widget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(main_window)

        # 定义动作
        self.btn_load_data.clicked.connect(self.load_data)
        self.btn_save_data.clicked.connect(self.save_page_data)
        self.btn_reset_data.clicked.connect(self.reset_data)
        self.btn_cut_data.clicked.connect(self.cut_data)
        self.btn_count_alpha.clicked.connect(self.calculate_alpha)
        self.btn_count_factor.clicked.connect(self.calculate_factor)
        self.btn_check_value.clicked.connect(self.check_value)

    def retranslateUi(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("main_window", "数据自动分析脚本v0.9.1"))
        self.btn_load_data.setText(_translate("main_window", "打开数据"))
        self.btn_save_data.setText(_translate("main_window", "保存数据"))
        self.btn_reset_data.setText(_translate("main_window", "重置数据"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.data_load), _translate("main_window", "数据加载"))
        self.lbl_tag_2_titile_1.setText(_translate("main_window", "处理结果"))
        self.lbl_tag_2_title_2.setText(_translate("main_window", "量表数据提取"))
        self.btn_cut_data.setText(_translate("main_window", "提取数据"))
        self.label_alpha.setText(_translate("main_window", "量表alpha计算"))
        self.btn_count_alpha.setText(_translate("main_window", "计算量表各项alpha值"))
        self.label_factor.setText(_translate("main_window", "探索性因子分析"))
        self.btn_count_factor.setText(_translate("main_window", "对数据进行因子分析"))
        self.label_check_value.setText(_translate("main_window", "检查量表数据"))
        self.btn_check_value.setText(_translate("main_window", "计算量表数据质量"))
        self.ckb_ignore_small.setText(_translate("main_window", "是否在结果中忽略小系数"))
        self.label_small_limit.setText(_translate("main_window", "小系数："))
        self.label_col_tocheck.setText(_translate("main_window", "请输入需要检测的量表标题（不含序号）"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.pre_test), _translate("main_window", "数据基础分析"))
        self.tab_widget.setTabText(self.tab_widget.indexOf(self.create_data), _translate("main_window", "模拟数据"))
        self.lbl_data_show.setText(_translate("main_window", "当前数据装载情况：未加载"))
        self.label.setText(_translate("main_window", "健二制作"))

    def show_error(self, e: Exception):
        QMessageBox.warning(self, '错误', f'出现{e}错误，请联系作者\n'
                                        f'dingjianyi0508@gmail.com')

    def load_data(self):
        fname = QFileDialog.getOpenFileName(self,
                                            caption='文件读取',
                                            directory='.',
                                            filter='')
        file_dir = fname[0]
        try:
            self._data = pd.read_excel(file_dir)
            ic(self._data.shape)
            self.show_data()
        except Exception as e:
            self.show_error(e)

    def show_give_data(self, data: pd.DataFrame):
        self.show_current_data.setColumnCount(data.shape[1])
        self.show_current_data.setRowCount(data.shape[0])
        self.show_current_data.setHorizontalHeaderLabels(data.columns)
        for row in range(data.shape[0]):
            for col in range(data.shape[1]):
                item = QtWidgets.QTableWidgetItem(str(data.iloc[row, col]))
                self.show_current_data.setItem(row, col, item)
        self.lbl_data_show.setText(f'当前数据为：{data.shape[0]}行，{data.shape[1]}列')

    def show_data(self):
        self.show_give_data(data=self._data)

    def show_slice_data(self):
        self._calculate_data = self._data.loc[:, self._all_questions]
        self.show_give_data(self._calculate_data)

    # 保存页面数据
    def save_page_data(self):
        try:
            self._save_dir = QFileDialog.getExistingDirectory(self)
            data_name = QInputDialog.getText(self, '请输入', '请输入保存文件名')
            col_num = self.show_current_data.columnCount()
            row_num = self.show_current_data.rowCount()
            empty_array = np.empty((row_num, col_num))
            self._temp_data = pd.DataFrame(data=empty_array)
            col_names = []
            for col in range(col_num):
                col_name = self.show_current_data.horizontalHeaderItem(col).text()
                col_names.append(col_name)
                for row in range(row_num):
                    self._temp_data.iloc[row, col] = self.show_current_data.item(row, col).text()
            self._temp_data.columns = col_names
            self._temp_data.to_excel(f'{self._save_dir}/{data_name[0]}.xlsx', index=None)
            ic(self._temp_data, self._temp_data.shape)
        except Exception as e:
            self.show_error(e)

    def reset_data(self):
        try:
            self.show_data()
        except Exception as e:
            self.show_error(e)

    def get_all_questions(self):
        full_list = []
        for name in self._col_names:
            full_list = full_list + self._col_dict[f"{name}"]
        return full_list

    def cut_data(self):
        try:
            ic(self._data)
            self._col_names, self._col_dict = agg_questions(self._data, mode=1)
            ic(self._col_names,self._col_dict)
            self._all_questions = self.get_all_questions()
            ic(self._all_questions)
            self.show_slice_data()
            self.tag_2_show_data.clear()
            self.tag_2_show_data.insertPlainText('\n 已提取数据。\n')
            self.tag_2_show_data.insertPlainText('\n————量表分类————\n\n')
            self.tag_2_show_data.insertPlainText(pprint_dict(self._col_dict))
        except Exception as e:
            #ic(e)
            self.show_error(e)

    def calculate_alpha(self):
        try:
            self._alpha_dict = agg_questions(self._calculate_data, mode=2)
            ic(self._col_dict)
            ic(self._alpha_dict)
            self.tag_2_show_data.clear()
            self.tag_2_show_data.insertPlainText('\n————量表分类————\n\n')
            self.tag_2_show_data.insertPlainText(pprint_dict(self._col_dict))
            self.tag_2_show_data.insertPlainText('\n')
            self.tag_2_show_data.insertPlainText('\n————各量表alpha值————\n\n')
            self.tag_2_show_data.insertPlainText(pprint_dict(self._alpha_dict))
        except Exception as e:
            self.show_error(e)

    def save_heatmap(self, matrix: np.array, save_dir: str):
        df_cm = pd.DataFrame(np.abs(matrix), index=self._calculate_data.columns)
        fig, ax = plt.subplots(figsize=(12, 10))
        sns.heatmap(df_cm, annot=True, cmap='BuPu', ax=ax)
        # 设置y轴字体的大小
        ax.tick_params(axis='x', labelsize=15)
        ax.set_title("Factor Analysis", fontsize=12)
        plt.savefig(save_dir + '/因子分析图.jpg')

    def calculate_factor(self):
        try:
            chi_square_value, p_value = calculate_bartlett_sphericity(self._calculate_data)
            kmo_all, kmo_model = calculate_kmo(self._calculate_data)
            self.tag_2_show_data.clear()
            self.tag_2_show_data.insertPlainText('\n————KMO与卡方检验结果————\n\n')
            self.tag_2_show_data.insertPlainText(f'卡方值为：{chi_square_value:.3f}\n卡方显著性为：{p_value:.3f}\n')
            self.tag_2_show_data.insertPlainText(f'KMO值为{kmo_model:.3f}\n')
            if kmo_model >= 0.6:
                num_factor, self._factor_matrix = factor_analyze(self._calculate_data)
                self._factor_matrix = abs(self._factor_matrix.round(3))
                self.tag_2_show_data.insertPlainText('\n————因子提取————\n\n')
                self.tag_2_show_data.insertPlainText(f'根据特征值大于1的规则，提取了{num_factor}个因子。\n')
                save_dir = QFileDialog.getExistingDirectory(self)
                factor_df = pd.DataFrame(self._factor_matrix, index=self._calculate_data.columns)
                if self.ckb_ignore_small.checkState() != 0:
                    self.tag_2_show_data.insertPlainText(f'\n执行忽略小系数。\n')
                    factor_df = factor_df.applymap(lambda x: x if x >= float(self.ety_small_limit.text()) else 0)
                factor_df.to_excel(save_dir + '/旋转后的因子矩阵.xlsx')
                self.save_heatmap(self._factor_matrix, save_dir)
                QMessageBox.information(self, '输出导出', '分析结果已导出至选择文件夹。')
            else:
                QMessageBox.warning(self, '数据质量不足', 'KMO值不足0.6,数据不适合使用因子分析。')
        except Exception as e:
            self.show_error(e)

    def check_value(self):
        try:
            to_exact = self._col_dict[f'{self.ety_col_check.text()}']
            self._temp_data = self._data.loc[:, to_exact]
            self._temp_data['标准差'] = self._temp_data.var(axis=1).round(3)
            self._temp_data['标准差排序'] = self._temp_data.标准差.rank(0).round()
            self.show_give_data(self._temp_data)
            QMessageBox.information(self, '计算完成', '请返回前页查看结果。')
        except Exception as e:
            self.show_error(e)


if __name__ == '__main__':
    QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myshow = Ui_main_window()
    myshow.show()
    sys.exit(app.exec_())