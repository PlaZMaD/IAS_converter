# -*- coding: utf-8 -*-
# pyinstaller.exe --onefile --hidden-import openpyxl.cell._writer --windowed --add-data "locale;locale"

from PyQt5.QtWidgets import QApplication, QPushButton, QFileDialog, QMainWindow, QTextBrowser, QWidget, QVBoxLayout, QTextEdit
from PyQt5 import QtCore, QtWidgets

import pandas as pd
import numpy as np

import os

import logging
import re
import sys

from pandas.api.types import is_datetime64_any_dtype as is_datetime
from copy import deepcopy as copy

import gettext
# import openpyxl
# print(openpyxl.__version__)

# lang1 = gettext.translation('messages', 'locale', languages=['en'])
# lang2 = gettext.translation('messages', 'locale', languages=['ru'], fallback=True)

bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__))) # get the bundle dir if bundled or simply the __file__ dir if not bundled
locales_dir = os.path.abspath(os.path.join(bundle_dir, 'locale'))

lang2 = gettext.translation('messages', locales_dir, languages=['en'])
lang2.install()

logging.basicConfig(level=logging.INFO, filename="log.log", filemode="w", force=True, encoding='utf-8')
logging.info("START")

known_columns = ["ALB", "APAR", "CH4", "CO2", "CO2C13", "D_SNOW", "DBH", "EVI", "FC", "FC_CMB", "FC_SSITC_TEST", "FCH4", "FCH4_CMB", "FCH4_PI", "FETCH_70", "FETCH_80", "FETCH_90", "FETCH_FILTER", "FETCH_MAX", "FH2O", "FN2O", "FN2O_CMB", "FNO", "FNO_CMB", "FNO2", "FNO2_CMB", "FO3", "FO3_SSITC_ TEST", "G", "GPP_PI", "H", "H_PI", "H_SSITC_TEST", "H2O", "LE", "LE_PI", "LE_SSITC_TEST", "LEAF_WET", "LW_IN", "LW_OUT", "MCRI", "MO_LENGTH", "MTCI", "N2O", "NDVI", "NEE_PI", "NETRAD", "NIRV", "NO", "NO2", "O3", "P", "P_RAIN", "P_SNOW", "PA", "PBLH", "PPFD_BC_IN", "PPFD_DIF", "PPFD_DIR", "PPFD_IN", "PPFD_OUT", "PRI", "R_UVA", "R_UVB", "RECO_PI", "REDCl", "REP", "RH", "RUNOFF", "SAP_DT", "SAP_FLOW", "SB", "SC", "SCH4", "SG", "SH", "SLE", "SPEC_NIR_IN", "SPEC_NIR_ OUT", "SPEC_PRI_ REF_IN", "SPEC_PRI_ REF_OUT", "SPEC_PRI_TGT_IN", "SPEC_PRI_TGT_OUT", "SPEC_RED_ IN", "SPEC_RED_ OUT", "SR", "STEMFLOW", "SW_BC_IN", "SW_DIF", "SW_IN", "SW_OUT", "SWC", "T_BOLE", "T_CANOPY", "T_DP", "T_SONIC", "T_SONIC_SIGMA", "TA", "TAU", "TAU_SSITC_TEST", "TCARI", "THROUGHFALL", "TS", "U_SIGMA", "USTAR", "V_SIGMA", "VPD_PI", "W_SIGMA", "WD", "WD_SIGMA", "WS", "WS_MAX", "WTD", "ZL", "CO2_STR", "H20_STR", "CH4_RSSI"]


def get_freq(df, time):
    try_max = 100
    try_ind = 0
    t_shift = 5
    start = 1
    deltas = df[time] - df[time].shift(1)
    while try_ind < try_max:
        del_arr = deltas.iloc[start + try_ind*t_shift : start + try_ind*t_shift + t_shift].values
        if not np.all(del_arr == del_arr[0]) and del_arr[0] is not None:
            try_ind = try_ind + 1
            continue
        else:
            return del_arr[0]


def column_checker(col_list):
  error_flag = 0

  #Проверка на наличие временнЫх колонок
  minimum_setup = ['TIMESTAMP_START', 'TIMESTAMP_END']
  if not set(col_list[:2]).issubset(minimum_setup):
    logging.error("DateTime columns ['TIMESTAMP_START', 'TIMESTAMP_END'] should go first. Check columns order and names.")
    error_flag = 1
  try:
    dtime_pos = col_list.index('DTime')
    if dtime_pos != 2:
      logging.error('Wrong DTime column position, should go 3!')
      error_flag = 1
  except ValueError:
    logging.info("No DTime column in the data")


  #Проверка на наличие дублирующихся названий
  seen = set()
  dupes = []
  for col in col_list:
      if col in seen:
          dupes.append(col)
      else:
          seen.add(col)
  if len(dupes) > 0:
    logging.error(f"There are columns with the same names: {dupes}")
    error_flag = 1


  #Сортировка по индексам
  full_col_list = copy(col_list)
  full_col_list = [f for f in full_col_list if f not in['TIMESTAMP_START', 'TIMESTAMP_END', 'DTime']]

  re_check = [re.fullmatch(r".*_[0-9]_[0-9]_[0-9]",f) for f in full_col_list]
  re_check = [f.string for f in re_check if f is not None]

  no_index_cols = [f for f in full_col_list if f not in re_check]
  if len(no_index_cols) > 0:
    error_flag = 1
    logging.error(f"Columns naming problem, columns {no_index_cols} do not have correct suffix structure.")

  unique_cols = list(set([f[:-6] for f in re_check]))
  for col in unique_cols:

    if col not in known_columns:
      logging.warning(f"The column {col} is not in the known columns list, please double-check the name!")

    col_with_inds = [f for f in re_check if re.fullmatch(f"{col}_[0-9]_[0-9]_[0-9]", f)]
    if len(col_with_inds)==1:
      if col_with_inds[0][-5:] != '1_1_1':
        logging.error(f"Suffix problem for {col_with_inds}, 1_1_1 should be used in case of a single variable.")
        error_flag = 1
    else:
      index_structure = np.array([f[-5:].split('_') for f in col_with_inds]).astype(int)
      for suf in range(3):
        ind_line = index_structure[:, suf]
        if ind_line.min() != 1:
          logging.error(f'Check suffix for {col}, not starting from 1')
          error_flag = 1
        # if ind_line.max() > len(ind_line)+1:
        #   logging.error(f'Check suffix for {col}, incrementing by > 1')
        #   error_flag = 1

        prev = 1
        sorted_inds = np.sort(ind_line)
        for i in sorted_inds:
          if i - prev > 1:
            logging.error(f'Check suffix for {col} with #{suf+1} index "{i}", incrementing by > 1')
            error_flag = 1
          prev = i

  return error_flag




def check_time(data, time_in, check_year=True):
  data_in = data.copy()
  data_in['default_index'] = data.index + 2
  data_in.index = data_in[time_in]
  outflag = 0
  # data_freq = pd.to_timedelta(get_freq(data_in, time_in))

  logging.info(f"Checking time column for {time_in}")
  correct_column_type = is_datetime(data_in[time_in])
  if not correct_column_type:
    logging.info(f"{time_in} is not of correct type.")
    outflag = 1

  missed_values = data_in[time_in].isna()
  if missed_values.any():
    logging.error(f"Can't read timestamp. Please check entries near lines \n {data_in.loc[missed_values, 'default_index'].to_numpy()}")
    outflag = 1

  data_in_dup = data_in.drop(data_in[missed_values].index, axis=0)

  if check_year and (not (data_in_dup[time_in].dt.year.to_numpy()[0] == data_in_dup[time_in].dt.year.to_numpy()[:-1]).all()):
    years = data_in_dup[time_in].dt.year.unique()
    examples = [int(data_in_dup.query(f'{time_in}.dt.year=={f}')['default_index'].to_numpy()[0]) for f in years]
    logging.error(f"There should be only one year presented in file, got {years}, i.e. lines {examples}!")
    outflag = 1

  if data_in_dup.index.duplicated(keep='first').any():
      logging.error(f"Duplicated timestamps! check lines:{ data_in_dup.loc[data_in_dup.index.duplicated(), 'default_index'].to_numpy()}")
      outflag = 1

  return outflag

def final_time_check(data, time_in):
  data_in = data.copy()
  data_in['default_index'] = data.index + 2
  data_in.index = data_in[time_in]
  outflag = 0

  data_freq = pd.to_timedelta(get_freq(data_in, time_in))

  year = data_in.index.year.to_numpy()[0]
  if "TIMESTAMP_START" in time_in:
    start = f"{year}.01.01 00:00"
    end = f"{year}.12.31 23:30"
  else:
    if "TIMESTAMP_END" in time_in:
      start = f"{year}.01.01 00:30"
      end = f"{year+1}.01.01 00:00"
    else:
      logging.error("Strange time column name!")
      return 1
  correct_index = pd.date_range(start=pd.to_datetime(start, format="%Y.%m.%d %H:%M"), end=pd.to_datetime(end, format="%Y.%m.%d %H:%M"),
                                                freq=pd.to_timedelta(data_freq))
  extra_index = data_in.index.difference(correct_index)
  missing_index = correct_index.difference(data_in.index)
  if len(missing_index)>0:
    outflag = 1
    logging.error(f"Missing values: {missing_index.astype(str)}")
  if len(extra_index)>0:
    outflag = 1
    logging.error(f"Extra values: {extra_index.astype(str)}, lines: {data_in.loc[extra_index, 'default_index'].to_numpy()}")

  logging.info("\n")

  return outflag



def check_file(path_to_file):
    file_name = path_to_file
    something, file_extension = os.path.splitext(file_name)

    ext = 'csv' if '.csv' in file_extension.lower() else None
    ext = 'excel' if file_extension.lower() in ['.xls', '.xlsx'] else ext

    load_func = None
    l_config = {}

    if ext == 'csv':
        load_func = pd.read_csv
        # l_config['sep'] = sep

    elif ext == 'excel':
        load_func = pd.read_excel
        l_config['sheet_name'] = None

    else:
        logging.error(_("Select CSV, XLS or XLSX file."))
        return 2

    try:
        data = load_func(file_name, **l_config)
    except Exception as e:
        logging.error(e)
        assert False

    if ext == 'excel':
        if len(data) > 1:
            logging.error("Several lists in data file!")
            assert False
        data = next(iter(data.values()))
        logging.info(f"File {file_name} loaded.\n")
    total_errors = 0
    columns = list(copy(data.columns))
    if len(columns)<3:
        logging.error("Not enough columns, please check your file. If you are using csv - make sure that it uses comma as a separator.")
        return 0

    check_column = column_checker(columns)
    total_errors = total_errors + check_column

    time_checks = []
    final_time_checks = []

    for col in ['TIMESTAMP_START', 'TIMESTAMP_END']:
        if col not in columns:
            time_checks.append(1)
            logging.error(f'No column {col} in the file!')
            continue

        data[f'{col}_datetime'] = pd.to_datetime(data[col].fillna(0).astype(int), format="%Y%m%d%H%M", errors='coerce')
        time_checks.append(check_time(data, f"{col}_datetime"))
        if not time_checks[-1]:
            final_time_checks.append(final_time_check(data, f"{col}_datetime"))
        else:
            final_time_checks.append(1)
            logging.warning(f"No  final timestamp checks for {col}, please fix all errors to complete this step.")

    total_errors = total_errors + np.sum(time_checks) + np.sum(final_time_checks)
    col_errors = 0

    for col in columns:
        if col in ['TIMESTAMP_START', 'TIMESTAMP_END'] or col not in data.columns:
            continue

        na_test = data[col].isna().all()
        if na_test:
            col_errors = col_errors + 1
            logging.error(f"The column {col} is empty.")
            continue

        inf_vals = data.loc[np.logical_or(data[col] == np.inf, data[col] == -np.inf)].index.to_numpy()
        if len(inf_vals) > 0:
            logging.error(f"INF val in {col} at the position {inf_vals}")
            col_errors = col_errors + 1

        check_col = pd.to_numeric(data[col].fillna(-9999), errors='coerce')

        inf_vals = check_col.loc[np.logical_or(check_col == np.inf, check_col == -np.inf)].index.to_numpy()
        if len(inf_vals) > 0:
            logging.error(f"INF val in {col} at the position {inf_vals}")
            col_errors = col_errors + 1

        wrong_vals = check_col.loc[check_col.isna()].index + 2
        if len(wrong_vals) > 0:
            col_errors = col_errors + 1
            logging.error(f"Non numerical values in {col} at lines {wrong_vals.to_numpy()}")

    total_errors = total_errors + col_errors
    print(f"{total_errors} errors in total, check logs!")

    return total_errors


class MyHandler(logging.StreamHandler, QtCore.QObject):
    # Сигнал, передающий отформатированное сообщение логгера
    # по сути мы его можем теперь ловить любым виджетом
    # и что угодно с ним делать, хоть на QLabel выводить
    message = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        logging.StreamHandler.__init__(self)
        QtCore.QObject.__init__(self, parent)

    def emit(self, record: str):
        # record - строка которую мы передали логгеру
        # метод формат применяет к ней форматтер если он установлен
        log_msg = self.format(record)

        # Раскрасим строку перед отправкой
        if "DEBUG" in log_msg:
            text = f"""<span style='color:#888888;'>{log_msg}</span>"""
        elif "INFO" in log_msg:
            text = f"""<span style='color:#008800;'>{log_msg}</span>"""
        elif "WARNING" in log_msg:
            text = f"""<span style='color:#888800;'>{log_msg}</span>"""
        elif "ERROR" in log_msg:
            text = f"""<span style='color:#880000;'>{log_msg}</span>"""
        elif "CRITICAL" in log_msg:
            text = f"""<span style='color:#000088;'>{log_msg}</span>"""

        # генерируем сигнал, передаем раскрашенный текст
        self.message.emit(text)

        # сбрасываем буфер
        self.flush()

# class LogWidget(QtWidgets.QTextEdit):
#     def __init__(self, handler: logging.StreamHandler, parent=None):
#         QtWidgets.QTabWidget.__init__(self, parent)
#         # Сохраним хендлер чтобы его сборщик мусора не прибил
#         self.handler = handler
#
#         # Конектим сигнал от хендлера
#         self.handler.message.connect(self.__updateText)
#
#     def __scrollDown(self):
#         logger.debug(f"{self.__class__.__name__}.__scrollDown()")
#         scroll_bar = self.verticalScrollBar()
#         end_text = scroll_bar.maximum()
#         scroll_bar.setValue(end_text)
#
#     def __updateText(self, msg: str):
#         logger.debug(f"{self.__class__.__name__}.__updateText(msg)")
#         self.append(msg)
#         self.__scrollDown()



class MyApp(QWidget):
   def __init__(self, handler: logging.StreamHandler, parent=None):
        super().__init__()
        QtWidgets.QTabWidget.__init__(self, parent)

        self.button = QPushButton("Open File", self)
        self.button.clicked.connect(self.openFileDialog)

        self.tab = QTextEdit()
        self.tab.setAcceptRichText(True)

        self.handler = handler
        self.handler.message.connect(self.__updateText)

        vbox = QVBoxLayout()
        vbox.addWidget(self.button, 0)
        vbox.addWidget(self.tab, 1)

        self.setLayout(vbox)

        self.setWindowTitle('QTextBrowser')
        self.setGeometry(300, 300, 1000, 600)
        self.show()
        logging.info("READY")

   def openFileDialog(self):
    file_dialog = QFileDialog(self)
    file_dialog.setWindowTitle("Open File")
    file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
    file_dialog.setViewMode(QFileDialog.ViewMode.Detail)

    if file_dialog.exec():
        selected_files = file_dialog.selectedFiles()
        logging.info(f"Selected File: {selected_files[0]}")
        check_file(selected_files[0])

   def __updateText(self, msg: str):
       logger.debug(f"{self.tab.__class__.__name__}.__updateText(msg)")
       self.tab.append(msg)


if __name__ == '__main__':

    logger = logging.getLogger()
    stream_formatter = logging.Formatter(
        "%(module)s: %(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(stream_formatter)
    stream_handler.setLevel(logging.INFO)


    logger.addHandler(stream_handler)

    # # И с файл хендлером для вывода лога в файл, уровень [DEBUG]
    # file_formatter = logging.Formatter(
    #     "%(module)s: %(asctime)s [%(levelname)s] %(message)s",
    #     datefmt="%Y-%m-%d %H:%M:%S",
    # )
    # file_path = "debug.log"
    # file_handler = logging.FileHandler(file_path, mode='w', encoding='utf-8')
    # file_handler.setLevel(logging.DEBUG)
    # file_handler.setFormatter(file_formatter)
    # logger.addHandler(file_handler)

    logger.critical(
        "Process init"
        )

    formatter = logging.Formatter(
        "[%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        )

    gui_stream_handler = MyHandler()  # наш хендлер, выдающий сигналы
    gui_stream_handler.setFormatter(formatter)
    # связываем его с общим для всего логгером
    logging.getLogger()
    logger.addHandler(gui_stream_handler)
    app = QApplication(sys.argv)
    ex = MyApp(gui_stream_handler)

    sys.exit(app.exec())