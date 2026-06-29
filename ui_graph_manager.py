# ui_graph_manager.py
import sqlite3
import pandas as pd
import os
from datetime import timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.ticker as ticker

from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QListWidget, QAbstractItemView, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
# 🌟 미리보기창(QPrintPreviewDialog) 및 관련 모듈 임포트
from PyQt5.QtGui import QPainter, QPageLayout 
from PyQt5.QtPrintSupport import QPrinter, QPrintPreviewDialog

import db_manager

# 그래프 내부에 한글(맑은 고딕)과 마이너스 부호가 깨지는 것을 방지합니다.
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

class GraphManager(QWidget):
    def __init__(self, parent_window):
        super().__init__()
        self.main_win = parent_window  # 메인 윈도우의 날짜창 등에 접근하기 위한 참조
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        
        # 상단 컨트롤 레이아웃
        graph_ctrl = QHBoxLayout()
        
        # [왼쪽 축] 다중 선택 리스트 위젯
        self.data_selector = QListWidget()
        self.data_selector.setSelectionMode(QAbstractItemView.MultiSelection) 
        self.data_selector.addItems(db_manager.DATA_LABELS)
        self.data_selector.setMaximumHeight(80) 
        
        # [오른쪽 보조축] 다중 선택 리스트 위젯
        self.right_axis_selector = QListWidget()
        self.right_axis_selector.setSelectionMode(QAbstractItemView.MultiSelection)
        self.right_axis_selector.setMaximumHeight(80) 

        self.period_selector = QComboBox()
        self.period_selector.addItems(["일간 (실시간 데이터)", "주간 (시간별 평균)", "월간 (시간별 평균)"])
        
        # 그래프 제어용 버튼 생성
        self.btn_save_graph = QPushButton("📊 그래프 저장")
        self.btn_save_graph.setStyleSheet("background-color: #34495e; color: white; font-weight: bold; min-height: 30px;")
        
        self.btn_print_graph = QPushButton("🖨️ 미리보기 및 인쇄") # 🌟 명칭 변경
        self.btn_print_graph.setStyleSheet("background-color: #16a085; color: white; font-weight: bold; min-height: 30px;")
        
        graph_ctrl.addWidget(QLabel("기본(왼쪽) 축 필드 선택:"))
        graph_ctrl.addWidget(self.data_selector)
        graph_ctrl.addWidget(QLabel("➡️ 오른쪽(보조) 축으로 보낼 필드:")) 
        graph_ctrl.addWidget(self.right_axis_selector)     
        graph_ctrl.addWidget(QLabel("조회 기간:"))
        graph_ctrl.addWidget(self.period_selector)
        graph_ctrl.addWidget(self.btn_save_graph)
        graph_ctrl.addWidget(self.btn_print_graph)
        layout.addLayout(graph_ctrl)

        # 캔버스 배치
        self.canvas = FigureCanvas(Figure(figsize=(8, 5)))
        self.ax = self.canvas.figure.add_subplot(111)
        layout.addWidget(self.canvas)

        # 내부 이벤트 연결
        self.period_selector.currentIndexChanged.connect(self.update_graph)
        self.data_selector.itemSelectionChanged.connect(self.sync_right_axis_list)
        self.right_axis_selector.itemSelectionChanged.connect(self.update_graph)
        self.btn_save_graph.clicked.connect(self.save_graph_to_pdf)
        self.btn_print_graph.clicked.connect(self.show_print_preview) # 🌟 미리보기 함수로 변경

        # 초기 신호 잠금 및 해제 제어
        self.data_selector.blockSignals(True)
        self.right_axis_selector.blockSignals(True)
        if self.data_selector.count() > 0:
            self.data_selector.item(0).setSelected(True)
        self.data_selector.blockSignals(False)
        self.right_axis_selector.blockSignals(False)

    def sync_right_axis_list(self):
        self.data_selector.blockSignals(True)
        self.right_axis_selector.blockSignals(True)
        try:
            prev_selected = [item.text() for item in self.right_axis_selector.selectedItems()]
            left_selected = [item.text() for item in self.data_selector.selectedItems()]
            
            self.right_axis_selector.clear()
            if left_selected:
                self.right_axis_selector.addItems(left_selected)
                for i in range(self.right_axis_selector.count()):
                    item = self.right_axis_selector.item(i)
                    if item.text() in prev_selected:
                        item.setSelected(True)
        finally:
            self.data_selector.blockSignals(False)
            self.right_axis_selector.blockSignals(False)
        self.update_graph()

    def update_graph(self):
        if self.main_win.stack.currentIndex() != 1: return
        
        selected_items = self.data_selector.selectedItems()
        if not selected_items:
            self.canvas.figure.clf()
            self.ax = self.canvas.figure.add_subplot(111)
            self.ax.text(0.5, 0.5, "비교할 필드를 선택해주세요.", ha='center')
            self.canvas.draw()
            return
            
        target_cols = [item.text() for item in selected_items]
        period = self.period_selector.currentText()
        selected_date = self.main_win.qdate.date().toPyDate() 

        right_cols = [item.text() for item in self.right_axis_selector.selectedItems()]
        right_cols = [col for col in right_cols if col in target_cols] 
        left_cols = [col for col in target_cols if col not in right_cols]

        cols_str = ', '.join([f'"{col}"' for col in target_cols])
        
        try:
            conn = sqlite3.connect(db_manager.DB_NAME)
            if "일간" in period:
                query = f'SELECT log_time, {cols_str} FROM raw_data WHERE log_date = ? ORDER BY log_time ASC'
                params = (selected_date.strftime('%Y-%m-%d'),)
            else:
                days = 7 if "주간" in period else 30
                start_date = selected_date - timedelta(days=days)
                query = f'SELECT log_date || \' \' || SUBSTR(log_time,1,5) as dt, {cols_str} FROM hourly_avg WHERE log_date BETWEEN ? AND ? ORDER BY log_date, log_time'
                params = (start_date.strftime('%Y-%m-%d'), selected_date.strftime('%Y-%m-%d'))

            df = pd.read_sql_query(query, conn, params=params)
            conn.close()
        except Exception as e:
            print(f"그래프 DB 조회 에러 발생: {e}")
            return

        self.canvas.figure.clf()
        self.ax = self.canvas.figure.add_subplot(111)
        
        if not df.empty:
            x_col = 'log_time' if "일간" in period else 'dt'
            all_lines = []
            
            # 🌟 선 색상 연속성 확보 핵심 로직
            # Matplotlib의 기본 10가지 색상 순서 배열을 가져옵니다.
            color_cycle = plt.rcParams['axes.prop_cycle'].by_key()['color']
            color_idx = 0
            
            # [1] 기본 축 - 왼쪽 렌더링 (실선 - 'o' 표식)
            for col in left_cols:
                if col in df.columns:
                    current_color = color_cycle[color_idx % len(color_cycle)]
                    line = self.ax.plot(df[x_col], df[col], marker='o', markersize=2, 
                                        color=current_color, label=col)
                    all_lines += line
                    color_idx += 1 # 색상 인덱스 누적
            
            if left_cols:
                self.ax.set_ylabel(', '.join(left_cols[:2]) + ('...' if len(left_cols) > 2 else ''), color='#1f77b4', fontweight='bold')
                self.ax.tick_params(axis='y', labelcolor='#1f77b4')
            else:
                self.ax.yaxis.set_visible(False)

            # [2] 보조 축 - 오른쪽 다중 렌더링 (★ 왼쪽 선 색상 뒤를 이어서 매핑, 대조를 위해 점선 '--' 및 '^' 표식 사용)
            if right_cols:
                ax2 = self.ax.twinx()
                for col in right_cols:
                    if col in df.columns:
                        current_color = color_cycle[color_idx % len(color_cycle)]
                        line = ax2.plot(df[x_col], df[col], marker='^', markersize=3, linestyle='--', 
                                        color=current_color, label=f"{col} (우)")
                        all_lines += line
                        color_idx += 1 # 보조축도 색상 인덱스를 계속 이어 나감
                
                ax2.set_ylabel(', '.join(right_cols[:2]) + ('...' if len(right_cols) > 2 else ''), color='#ff7f0e', fontweight='bold')
                ax2.tick_params(axis='y', labelcolor='#ff7f0e')
                ax2.grid(False) 

            # ====== 이중 축 마우스 먹통 문제를 해결한 최종 범례 로직 ======
            if all_lines:
                labels = [l.get_label() for l in all_lines]
                
                # [핵심] 보조축(ax2)이 생성되어 있다면 ax2의 권한으로 범례를 그리고, 
                # 보조축이 없다면 기본축(self.ax)의 권한으로 범례를 그립니다.
                if 'ax2' in locals():
                    leg = ax2.legend(all_lines, labels, loc='upper right')
                else:
                    leg = self.ax.legend(all_lines, labels, loc='upper right')
                
                # 범례 드래그 기능 활성화 (이제 무조건 작동합니다!)
                leg.set_draggable(True)
            # ============================================================
                
            self.ax.xaxis.set_major_locator(ticker.MaxNLocator(nbins=10))
            self.ax.set_title(f"선택 필드 {period} 분석 (양방향 멀티 축 제어)", fontsize=13, fontweight='bold')
            self.ax.grid(True, linestyle='--')
            self.canvas.figure.autofmt_xdate() 
        else:
            self.ax.text(0.5, 0.5, "데이터가 존재하지 않습니다.", ha='center')
        
        self.canvas.draw()

    # 🌟 [수정] 미리보기 창에서는 오직 '실제 인쇄(하드카피)'만 담당하도록 제안 제거
    def show_print_preview(self):
        """[미리보기 및 인쇄] 버튼: 실제 프린터로 종이 인쇄를 하기 위한 미리보기"""
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageOrientation(QPageLayout.Landscape) # 가로 방향
        
        # ❌ printer.setOutputFileName(...) 코드를 제거했습니다. 
        # 이제 자꾸 파일 저장으로 강제 전환되지 않고, 실제 프린터를 선택해 인쇄할 수 있습니다.
        
        preview_dialog = QPrintPreviewDialog(printer, self)
        preview_dialog.paintRequested.connect(self.render_print_page)
        preview_dialog.resize(1100, 800)
        preview_dialog.exec_()

    def render_print_page(self, printer):
        """미리보기 및 인쇄 장치에 그래프를 그려주는 렌더러"""
        try:
            painter = QPainter(printer)
            pixmap = self.canvas.grab() # 현재 그래프 화면 캡처
            
            rect = painter.viewport()
            size = pixmap.size()
            size.scale(rect.size(), Qt.KeepAspectRatio) # 용지 비율 맞춤
            
            painter.setViewport(rect.x(), rect.y(), size.width(), size.height())
            painter.setWindow(pixmap.rect())
            painter.drawPixmap(0, 0, pixmap)
            painter.end()
        except Exception as e:
            print(f"인쇄 렌더링 오류: {e}")

    # 🌟 [신규 추가] PDF 저장을 완벽하게 분리하여 기본 파일명까지 제안하는 함수
    def save_graph_to_pdf(self):
        """[PDF로 저장] 전용 함수: 과장님이 원하셨던 기본 파일명을 제안하며 파일로 저장"""
        selected_date = self.main_win.qdate.date().toString("yyyy-MM-dd")
        period = self.period_selector.currentText().split()[0]
        
        # 기본 파일명 제안 생성
        default_filename = f"{selected_date}_{period}_부하분석그래프.pdf"
        
        # 파일 저장 창 열기
        filepath, _ = QFileDialog.getSaveFileName(
            self, "그래프를 PDF로 저장", default_filename, "PDF 문서 (*.pdf)"
        )
        
        if filepath:
            try:
                printer = QPrinter(QPrinter.HighResolution)
                printer.setPageOrientation(QPageLayout.Landscape)
                printer.setOutputFormat(QPrinter.PdfFormat) # 👈 PDF 포맷 강제 지정
                printer.setOutputFileName(filepath)         # 👈 지정한 경로로 파일명 주입
                
                # PDF 파일에 그래프 그려서 내보내기
                self.render_print_page(printer)
                
                QMessageBox.information(self, "저장 완료", f"그래프가 PDF 파일로 안전하게 저장되었습니다.\n\n경로: {filepath}")
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", f"PDF 저장 중 오류가 발생했습니다:\n{e}")