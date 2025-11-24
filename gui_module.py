# Interface Gráfica

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QWidget,
    QStackedWidget, QLineEdit, QLabel, QVBoxLayout, QSizePolicy,
    QMessageBox, QTabWidget, QFormLayout,
    QTableWidget, QTableWidgetItem, QHBoxLayout, QFileDialog, QDialog,
    QComboBox, QAbstractItemView, QCheckBox, QSpinBox, QHeaderView
)
from PySide6.QtCore import Qt
from datetime import datetime
from db_module import insert_order, get_orders, update_order

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdfcanvas
except Exception:
    pdfcanvas = None

try:
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    import matplotlib.pyplot as plt
except Exception:
    FigureCanvas = None
    plt = None

app = QApplication()

# Tema de cores da interface
LIGHT_BLUE_BG = "background-color: #456882; color: white;"
BUTTON_STYLE = (
    "background-color: #456882; color: white; font-weight: bold;"
    " font-size: 18px; padding: 8px 14px; border-radius: 10px;"
)
COMMON_MIN_HEIGHT = 54


# Tela de Registrar Pedido
class TelaRegistrarPedido(QWidget):
    def __init__(self):
        super().__init__()

        vbox = QVBoxLayout()
        vbox.setAlignment(Qt.AlignTop)
        self.setLayout(vbox)
        self.setStyleSheet("color: white;")

        self.main_title = QLabel("Registrar Pedido")
        self.main_title.setStyleSheet("font-weight: bold; font-size: 32px; color: white;")
        self.main_title.setAlignment(Qt.AlignCenter)
        vbox.addWidget(self.main_title)
        vbox.addSpacing(12)

        container = QWidget()
        container.setFixedWidth(920)
        container_layout = QVBoxLayout()
        container.setLayout(container_layout)
        vbox.addWidget(container, 0, Qt.AlignHCenter)

        self.tabs = QTabWidget()
        self.tabs.tabBar().setStyleSheet("QTabBar::tab { color: white; }")
        container_layout.addWidget(self.tabs)

        self.tab_dados = QWidget()
        form_dados = QFormLayout()
        form_dados.setLabelAlignment(Qt.AlignVCenter)
        self.tab_dados.setLayout(form_dados)

        self.input_name = QLineEdit()
        self.input_date = QLineEdit()
        self.input_date.setPlaceholderText('DD/MM/AAAA')
        self.input_address = QLineEdit()
        self.input_phone = QLineEdit()
        self.input_type = QLineEdit()
        self.input_description = QLineEdit()
        self.input_description.setPlaceholderText('ex: projeto/infraestrutura/drenagem')
        self.input_status = QLineEdit()
        self.input_status.setPlaceholderText('pendente / finalizado')

        inputs = [
            self.input_name, self.input_date, self.input_address,
            self.input_phone, self.input_type, self.input_description,
            self.input_status
        ]
        for fld in inputs:
            fld.setStyleSheet(LIGHT_BLUE_BG)
            fld.setMinimumHeight(COMMON_MIN_HEIGHT)

        def lbl(text):
            l = QLabel(text)
            l.setStyleSheet("color: white;")
            return l

        form_dados.addRow(lbl('<b>Nome do Cliente</b>'), self.input_name)
        form_dados.addRow(lbl('<b>Data</b>'), self.input_date)
        form_dados.addRow(lbl('<b>Endereço</b>'), self.input_address)
        form_dados.addRow(lbl('<b>Telefone</b>'), self.input_phone)
        form_dados.addRow(lbl('<b>Tipo do Pedido</b>'), self.input_type)
        form_dados.addRow(lbl('<b>Descrição</b>'), self.input_description)
        form_dados.addRow(lbl('<b>Status</b>'), self.input_status)

        self.tabs.addTab(self.tab_dados, 'Dados')

        # Aba Valor
        self.tab_valor = QWidget()
        form_val = QFormLayout()
        self.tab_valor.setLayout(form_val)

        cb_container = QWidget()
        cb_layout = QHBoxLayout()
        cb_container.setLayout(cb_layout)
        self.cb_avista = QCheckBox("À vista")
        self.cb_parcelado = QCheckBox("Parcelado")
        self.cb_avista.setStyleSheet("color: white;")
        self.cb_parcelado.setStyleSheet("color: white;")
        cb_layout.addWidget(self.cb_avista)
        cb_layout.addWidget(self.cb_parcelado)
        self.cb_avista.toggled.connect(lambda val: self._cb_exclusive(self.cb_avista, self.cb_parcelado))
        self.cb_parcelado.toggled.connect(lambda val: self._on_parcelado_toggled(val))

        self.input_valor_total = QLineEdit()
        self.input_valor_total.setPlaceholderText('Valor total (apenas números)')
        self.input_valor_total.setStyleSheet(LIGHT_BLUE_BG)
        self.input_valor_total.setMinimumHeight(COMMON_MIN_HEIGHT)

        self.input_parcelas = QSpinBox()
        self.input_parcelas.setMinimum(1)
        self.input_parcelas.setMaximum(120)
        self.input_parcelas.setValue(1)
        self.input_parcelas.setEnabled(False)
        self.input_parcelas.setStyleSheet(LIGHT_BLUE_BG)
        self.input_parcelas.setMinimumHeight(COMMON_MIN_HEIGHT)

        form_val.addRow(lbl('<b>Forma de Pagamento</b>'), cb_container)
        form_val.addRow(lbl('<b>Valor Total</b>'), self.input_valor_total)
        form_val.addRow(lbl('<b>Parcelas</b>'), self.input_parcelas)

        self.tab_valor.setStyleSheet("color: white;")
        self.tabs.addTab(self.tab_valor, 'Valor')

        # Aba Parceiros
        self.tab_parceiros = QWidget()
        parceiros_layout = QVBoxLayout()
        self.tab_parceiros.setLayout(parceiros_layout)

        add_layout = QFormLayout()
        self.p_nome = QLineEdit()
        self.p_telefone = QLineEdit()
        self.p_valor = QLineEdit()
        self.p_data_pag = QLineEdit()
        self.p_data_pag.setPlaceholderText('DD/MM/AAAA')

        for fld in (self.p_nome, self.p_telefone, self.p_valor, self.p_data_pag):
            fld.setStyleSheet(LIGHT_BLUE_BG)
            fld.setMinimumHeight(COMMON_MIN_HEIGHT)

        add_layout.addRow('Nome', self.p_nome)
        add_layout.addRow('Telefone', self.p_telefone)
        add_layout.addRow('Valor', self.p_valor)
        add_layout.addRow('Data Pagamento', self.p_data_pag)

        btn_add_parceiro = QPushButton('Adicionar Parceiro')
        btn_add_parceiro.setStyleSheet(BUTTON_STYLE)
        btn_add_parceiro.setMinimumHeight(COMMON_MIN_HEIGHT)
        btn_add_parceiro.clicked.connect(self.adicionar_parceiro)
        parceiros_layout.addLayout(add_layout)
        parceiros_layout.addWidget(btn_add_parceiro)

        self.table_parceiros = QTableWidget()
        self.table_parceiros.setColumnCount(4)
        self.table_parceiros.setHorizontalHeaderLabels(['Nome', 'Telefone', 'Valor', 'Data Pagamento'])
        self.table_parceiros.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        parceiros_layout.addWidget(self.table_parceiros)

        btn_remove_parceiro = QPushButton('Remover Parceiro Selecionado')
        btn_remove_parceiro.setStyleSheet(BUTTON_STYLE)
        btn_remove_parceiro.setMinimumHeight(COMMON_MIN_HEIGHT)
        btn_remove_parceiro.clicked.connect(self.remover_parceiro)
        parceiros_layout.addWidget(btn_remove_parceiro)

        self.parceiros_lista = []
        self.tabs.addTab(self.tab_parceiros, 'Parceiros')

        self.btn_salvar = QPushButton('Registrar Pedido')
        self.btn_salvar.setStyleSheet(BUTTON_STYLE)
        self.btn_salvar.setMinimumHeight(COMMON_MIN_HEIGHT)
        self.btn_salvar.clicked.connect(self.registrar_pedido)
        container_layout.addWidget(self.btn_salvar)

    def _cb_exclusive(self, this, other):
        if this.isChecked():
            other.setChecked(False)
        if this is self.cb_avista and this.isChecked():
            self.input_parcelas.setValue(1)
            self.input_parcelas.setEnabled(False)

    def _on_parcelado_toggled(self, val):
        if val:
            if self.cb_avista.isChecked():
                self.cb_avista.setChecked(False)
            self.input_parcelas.setEnabled(True)
        else:
            self.input_parcelas.setValue(1)
            self.input_parcelas.setEnabled(False)

    def adicionar_parceiro(self):
        nome = self.p_nome.text().strip()
        telefone = self.p_telefone.text().strip()
        valor = self.p_valor.text().strip()
        data_pag = self.p_data_pag.text().strip()

        if not nome:
            QMessageBox.warning(self, 'Atenção', 'Informe o nome do parceiro')
            return
        try:
            valor_num = float(valor) if valor else 0
        except Exception:
            QMessageBox.warning(self, 'Atenção', 'Valor do parceiro inválido')
            return

        parceiro = {'nome': nome, 'telefone': telefone, 'valor': valor_num, 'data_pagamento': data_pag}
        self.parceiros_lista.append(parceiro)
        self._refresh_parceiros_table()
        self.p_nome.clear(); self.p_telefone.clear(); self.p_valor.clear(); self.p_data_pag.clear()

    def _refresh_parceiros_table(self):
        self.table_parceiros.setRowCount(len(self.parceiros_lista))
        for i, p in enumerate(self.parceiros_lista):
            self.table_parceiros.setItem(i, 0, QTableWidgetItem(p.get('nome', '')))
            self.table_parceiros.setItem(i, 1, QTableWidgetItem(p.get('telefone', '')))
            self.table_parceiros.setItem(i, 2, QTableWidgetItem(str(p.get('valor', ''))))
            self.table_parceiros.setItem(i, 3, QTableWidgetItem(p.get('data_pagamento', '')))

    def remover_parceiro(self):
        row = self.table_parceiros.currentRow()
        if row >= 0 and row < len(self.parceiros_lista):
            del self.parceiros_lista[row]
            self._refresh_parceiros_table()

    def registrar_pedido(self):
        try:
            valor_total = float(self.input_valor_total.text().strip() or 0)
        except Exception:
            QMessageBox.warning(self, 'Atenção', 'Valor total inválido. Use apenas números.')
            return

        forma = 'avista' if self.cb_avista.isChecked() else ('parcelado' if self.cb_parcelado.isChecked() else 'avista')
        parcelas = int(self.input_parcelas.value()) if forma == 'parcelado' else 1

        descricao = self.input_description.text().strip()
        categoria = subcategoria = especificacao = ''
        if descricao:
            parts = [p.strip() for p in descricao.split('/') if p.strip()]
            if len(parts) >= 1:
                categoria = parts[0]
            if len(parts) >= 2:
                subcategoria = parts[1]
            if len(parts) >= 3:
                especificacao = parts[2]

        dados = {
            'cliente': self.input_name.text().strip(),
            'data': self.input_date.text().strip(),
            'endereco': self.input_address.text().strip(),
            'telefone': self.input_phone.text().strip(),
            'tipo_pedido': self.input_type.text().strip(),
            'descricao': descricao,
            'categoria': categoria,
            'subcategoria': subcategoria,
            'especificacao': especificacao,
            'status': self.input_status.text().strip() or 'pendente',
            'valor': {'tipo': forma, 'valor_total': valor_total, 'parcelas': parcelas},
            'parceiros': {'tipo': 'com_parceiros' if len(self.parceiros_lista) > 0 else 'sem_parceiros', 'lista': self.parceiros_lista}
        }

        for p in dados['parceiros']['lista']:
            if isinstance(p.get('data_pagamento'), str) and p.get('data_pagamento'):
                try:
                    p['data_pagamento'] = datetime.strptime(p['data_pagamento'], '%d/%m/%Y')
                except Exception:
                    p['data_pagamento'] = None

        try:
            inserted_id = insert_order(dados)
        except Exception as e:
            QMessageBox.critical(self, 'Erro', f'Falha ao inserir pedido: {e}')
            return

        QMessageBox.information(self, 'Sucesso', f'Pedido registrado com ID: {inserted_id}')
        self.input_name.clear(); self.input_date.clear(); self.input_address.clear(); self.input_phone.clear()
        self.input_type.clear(); self.input_description.clear(); self.input_status.clear()
        self.cb_avista.setChecked(False); self.cb_parcelado.setChecked(False)
        self.input_valor_total.clear()
        self.input_parcelas.setValue(1); self.input_parcelas.setEnabled(False)
        self.parceiros_lista = []
        self._refresh_parceiros_table()

#  Janela de Editar Pedido
class EditPedidoDialog(QDialog):
    def __init__(self, pedido_doc: dict, parent=None):
        super().__init__(parent)
        self.pedido_doc = pedido_doc
        self.setWindowTitle(f"Editar Pedido — {pedido_doc.get('pedido')}")
        self.resize(720, 520)

        layout = QVBoxLayout()
        self.setLayout(layout)
        self.setStyleSheet("color: white;")

        form = QFormLayout()
        self.e_cliente = QLineEdit(pedido_doc.get('cliente', ''))
        data_val = pedido_doc.get('data')
        self.e_data = QLineEdit(data_val.strftime('%d/%m/%Y') if hasattr(data_val, 'strftime') else (data_val or ''))
        self.e_endereco = QLineEdit(pedido_doc.get('endereco', ''))
        self.e_telefone = QLineEdit(pedido_doc.get('telefone', ''))
        self.e_tipo = QLineEdit(pedido_doc.get('tipo_pedido', ''))
        descricao_val = pedido_doc.get('descricao')
        if not descricao_val:
            parts = []
            if pedido_doc.get('categoria'):
                parts.append(pedido_doc.get('categoria'))
            if pedido_doc.get('subcategoria'):
                parts.append(pedido_doc.get('subcategoria'))
            if pedido_doc.get('especificacao'):
                parts.append(pedido_doc.get('especificacao'))
            descricao_val = '/'.join(parts)
        self.e_descricao = QLineEdit(descricao_val or '')
        self.e_status = QLineEdit(pedido_doc.get('status', 'pendente'))

        valor = pedido_doc.get('valor', {})
        self.e_valor_tipo = QLineEdit(valor.get('tipo', 'avista'))
        parcelas_val = valor.get('parcelas', 1) if isinstance(valor, dict) else 1
        self.e_parcelas = QSpinBox()
        self.e_parcelas.setMinimum(1)
        self.e_parcelas.setMaximum(120)
        self.e_parcelas.setValue(int(parcelas_val))
        self.e_valor_total = QLineEdit(str(valor.get('valor_total', 0)))

        for fld in (self.e_cliente, self.e_data, self.e_endereco, self.e_telefone,
                    self.e_tipo, self.e_descricao, self.e_status, self.e_valor_tipo, self.e_valor_total):
            fld.setStyleSheet(LIGHT_BLUE_BG)
            fld.setMinimumHeight(COMMON_MIN_HEIGHT)
        self.e_parcelas.setStyleSheet(LIGHT_BLUE_BG)
        self.e_parcelas.setMinimumHeight(COMMON_MIN_HEIGHT)

        form.addRow("Cliente", self.e_cliente)
        form.addRow("Data (DD/MM/AAAA)", self.e_data)
        form.addRow("Endereço", self.e_endereco)
        form.addRow("Telefone", self.e_telefone)
        form.addRow("Tipo", self.e_tipo)
        form.addRow("Descrição", self.e_descricao)
        form.addRow("Status", self.e_status)
        form.addRow("Valor Tipo", self.e_valor_tipo)
        form.addRow("Valor Total", self.e_valor_total)
        form.addRow("Parcelas", self.e_parcelas)

        layout.addLayout(form)

        btns = QHBoxLayout()
        self.btn_save = QPushButton("Salvar Alterações")
        self.btn_cancel = QPushButton("Cancelar")
        self.btn_save.setStyleSheet(BUTTON_STYLE)
        self.btn_cancel.setStyleSheet(BUTTON_STYLE)
        self.btn_save.setMinimumHeight(COMMON_MIN_HEIGHT)
        self.btn_cancel.setMinimumHeight(COMMON_MIN_HEIGHT)
        self.btn_save.clicked.connect(self.salvar)
        self.btn_cancel.clicked.connect(self.reject)
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_cancel)
        layout.addLayout(btns)

    def salvar(self):
        try:
            valor_total = float(self.e_valor_total.text().strip() or 0)
        except Exception:
            QMessageBox.warning(self, "Atenção", "Valor total inválido.")
            return

        descricao = self.e_descricao.text().strip()
        categoria = subcategoria = especificacao = ''
        if descricao:
            parts = [p.strip() for p in descricao.split('/') if p.strip()]
            if len(parts) >= 1:
                categoria = parts[0]
            if len(parts) >= 2:
                subcategoria = parts[1]
            if len(parts) >= 3:
                especificacao = parts[2]

        updates = {
            'cliente': self.e_cliente.text().strip(),
            'data': self.e_data.text().strip(),
            'endereco': self.e_endereco.text().strip(),
            'telefone': self.e_telefone.text().strip(),
            'tipo_pedido': self.e_tipo.text().strip(),
            'descricao': descricao,
            'categoria': categoria,
            'subcategoria': subcategoria,
            'especificacao': especificacao,
            'status': self.e_status.text().strip(),
            'valor': {
                'tipo': self.e_valor_tipo.text().strip() or 'avista',
                'valor_total': valor_total,
                'parcelas': int(self.e_parcelas.value() or 1)
            }
        }

        try:
            res = update_order(self.pedido_doc.get('pedido'), updates)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao atualizar: {e}")
            return

        if isinstance(res, dict) and res.get('matched', 0) == 0:
            QMessageBox.warning(self, "Aviso", "Pedido não encontrado para atualizar.")
        else:
            QMessageBox.information(self, "Sucesso", "Pedido atualizado.")
            self.accept()

#  Tela de Buscar Pedido
class TelaBuscarPedido(QWidget):
    def __init__(self):
        super().__init__()
        self.last_results = []

        vbox = QVBoxLayout()
        vbox.setAlignment(Qt.AlignTop)
        self.setLayout(vbox)
        self.setStyleSheet("color: white;")

        title = QLabel("Buscar Pedido")
        title.setStyleSheet("font-weight: bold; font-size: 28px; color: white;")
        title.setAlignment(Qt.AlignCenter)
        vbox.addWidget(title)
        vbox.addSpacing(8)

        search_row = QHBoxLayout()
        self.input_search = QLineEdit()
        self.input_search.setPlaceholderText('Digite qualquer termo (ex: drenagem, cliente, 2025-002, construtora)')
        self.date_from = QLineEdit(); self.date_from.setPlaceholderText('Data de (DD/MM/AAAA)')
        self.date_to = QLineEdit(); self.date_to.setPlaceholderText('Data até (DD/MM/AAAA)')
        self.select_status = QComboBox(); self.select_status.addItems(['todos', 'pendente', 'finalizado'])

        for fld in (self.input_search, self.date_from, self.date_to):
            fld.setStyleSheet(LIGHT_BLUE_BG)
            fld.setMinimumHeight(COMMON_MIN_HEIGHT)

        btn_search = QPushButton('Pesquisar')
        btn_search.setStyleSheet(BUTTON_STYLE)
        btn_search.setMinimumHeight(COMMON_MIN_HEIGHT)
        btn_search.clicked.connect(self.buscar)

        search_row.addWidget(self.input_search)
        search_row.addWidget(self.date_from)
        search_row.addWidget(self.date_to)
        search_row.addWidget(self.select_status)
        search_row.addWidget(btn_search)
        vbox.addLayout(search_row)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(['Pedido', 'Cliente', 'Descrição', 'Status', 'Valor Total', 'Parcelas'])
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        vbox.addWidget(self.table)

        actions = QHBoxLayout()
        self.btn_view = QPushButton("Ver/Editar Selecionado"); self.btn_view.clicked.connect(self.edit_selected)
        self.btn_export_csv = QPushButton("Exportar CSV (selecionados)"); self.btn_export_csv.clicked.connect(lambda: self.export_selected('csv'))
        self.btn_export_xlsx = QPushButton("Exportar XLSX (selecionados)"); self.btn_export_xlsx.clicked.connect(lambda: self.export_selected('xlsx'))
        self.btn_export_pdf = QPushButton("Exportar PDF (selecionados)"); self.btn_export_pdf.clicked.connect(self.export_selected_pdf)
        self.btn_export_all = QPushButton("Exportar TODOS (CSV)"); self.btn_export_all.clicked.connect(lambda: self.export_all('csv'))

        for b in [self.btn_view, self.btn_export_csv, self.btn_export_xlsx, self.btn_export_pdf, self.btn_export_all]:
            b.setStyleSheet(BUTTON_STYLE)
            b.setMinimumHeight(COMMON_MIN_HEIGHT)
            actions.addWidget(b)

        vbox.addLayout(actions)

    def buscar(self):
        termo = self.input_search.text().strip()
        date_from = self.date_from.text().strip()
        date_to = self.date_to.text().strip()
        status = self.select_status.currentText()

        filtro = {}

        if termo:
            filtro['$or'] = [
                {"cliente": {"$regex": termo, "$options": "i"}},
                {"pedido": {"$regex": termo, "$options": "i"}},
                {"descricao": {"$regex": termo, "$options": "i"}},
                {"categoria": {"$regex": termo, "$options": "i"}},
                {"subcategoria": {"$regex": termo, "$options": "i"}},
                {"especificacao": {"$regex": termo, "$options": "i"}},
                {"endereco": {"$regex": termo, "$options": "i"}},
                {"tipo_pedido": {"$regex": termo, "$options": "i"}},
                {"parceiros.lista.nome": {"$regex": termo, "$options": "i"}},
                {"parceiros.lista.telefone": {"$regex": termo, "$options": "i"}}
            ]

        if date_from:
            filtro['date_from'] = date_from
        if date_to:
            filtro['date_to'] = date_to

        if status and status != 'todos':
            if '$or' in filtro:
                filtro = {'$and': [filtro, {'status': status}]}
            else:
                filtro['status'] = status

        resultados = get_orders(filtro)
        self.last_results = resultados
        self._fill_table()

    def _fill_table(self):
        self.table.setRowCount(len(self.last_results))
        for i, r in enumerate(self.last_results):
            pedido = r.get('pedido', '')
            cliente = r.get('cliente', '')
            descricao = r.get('descricao')
            if not descricao:
                parts = []
                if r.get('categoria'):
                    parts.append(r.get('categoria'))
                if r.get('subcategoria'):
                    parts.append(r.get('subcategoria'))
                if r.get('especificacao'):
                    parts.append(r.get('especificacao'))
                descricao = '/'.join(parts)
            status = r.get('status', '')
            valor = r.get('valor', {}).get('valor_total', '')
            parcelas = r.get('valor', {}).get('parcelas', 1)
            self.table.setItem(i, 0, QTableWidgetItem(str(pedido)))
            self.table.setItem(i, 1, QTableWidgetItem(cliente))
            self.table.setItem(i, 2, QTableWidgetItem(descricao or ''))
            self.table.setItem(i, 3, QTableWidgetItem(status))
            self.table.setItem(i, 4, QTableWidgetItem(str(valor)))
            self.table.setItem(i, 5, QTableWidgetItem(str(parcelas)))
        self.table.resizeColumnsToContents()

    def get_selected_row(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.last_results):
            return None
        return self.last_results[row]

    def get_selected_docs(self):
        rows = sorted(set([idx.row() for idx in self.table.selectedIndexes()]))
        docs = []
        for r in rows:
            if r < len(self.last_results):
                docs.append(self.last_results[r])
        return docs

    def edit_selected(self):
        doc = self.get_selected_row()
        if not doc:
            QMessageBox.information(self, "Aviso", "Selecione uma linha para editar.")
            return
        dlg = EditPedidoDialog(doc, self)
        if dlg.exec():
            self.buscar()

    def export_selected(self, kind='csv'):
        docs = self.get_selected_docs()
        if not docs:
            QMessageBox.information(self, "Aviso", "Selecione ao menos um pedido para exportar.")
            return
        if pd is None:
            QMessageBox.warning(self, "Dependência", "pandas não está instalado. Instale com: pip install pandas openpyxl")
            return
        df = self._docs_to_dataframe(docs)
        path, _ = QFileDialog.getSaveFileName(self, "Salvar arquivo", filter="CSV (*.csv)" if kind=='csv' else "Excel (*.xlsx)")
        if not path:
            return
        try:
            if kind == 'csv':
                df.to_csv(path, index=False)
            else:
                df.to_excel(path, index=False)
            QMessageBox.information(self, "Sucesso", f"Exportado para {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar: {e}")

    def export_all(self, kind='csv'):
        docs = self.last_results
        if not docs:
            QMessageBox.information(self, "Aviso", "Nenhum resultado para exportar.")
            return
        if pd is None:
            QMessageBox.warning(self, "Dependência", "pandas não está instalado. Instale com: pip install pandas openpyxl")
            return
        df = self._docs_to_dataframe(docs)
        path, _ = QFileDialog.getSaveFileName(self, "Salvar arquivo", filter="CSV (*.csv)")
        if not path:
            return
        try:
            df.to_csv(path, index=False)
            QMessageBox.information(self, "Sucesso", f"Exportado para {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar: {e}")

    def export_selected_pdf(self):
        docs = self.get_selected_docs()
        if not docs:
            QMessageBox.information(self, "Aviso", "Selecione ao menos um pedido para exportar.")
            return
        if pdfcanvas is None:
            QMessageBox.warning(self, "Dependência", "reportlab não está instalado. Instale com: pip install reportlab")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF", filter="PDF (*.pdf)")
        if not path:
            return
        try:
            c = pdfcanvas.Canvas(path, pagesize=A4)
            w, h = A4
            y = h - 40
            for doc in docs:
                c.setFont("Helvetica-Bold", 12)
                c.drawString(40, y, f"PEDIDO: {doc.get('pedido')}")
                y -= 18
                c.setFont("Helvetica", 10)
                c.drawString(40, y, f"CLIENTE: {doc.get('cliente')}")
                y -= 14
                data = doc.get('data')
                data_str = data.strftime('%d/%m/%Y') if hasattr(data, 'strftime') else str(data)
                c.drawString(40, y, f"DATA: {data_str}")
                y -= 14
                c.drawString(40, y, f"ENDEREÇO: {doc.get('endereco')}")
                y -= 14
                c.drawString(40, y, f"STATUS: {doc.get('status')}")
                y -= 14
                valor = doc.get('valor', {})
                c.drawString(40, y, f"VALOR: {valor.get('valor_total')} ({valor.get('tipo')})")
                y -= 18
                c.drawString(40, y, f"PARCELAS: {valor.get('parcelas', 1)}")
                y -= 18
                descricao = doc.get('descricao') or '/'.join(filter(None, [doc.get('categoria'), doc.get('subcategoria'), doc.get('especificacao')]))
                c.drawString(40, y, f"DESCRIÇÃO: {descricao}")
                y -= 18
                parceiros = doc.get('parceiros', {}).get('lista', [])
                if parceiros:
                    for p in parceiros:
                        dp = p.get('data_pagamento')
                        if hasattr(dp, 'strftime'):
                            dp = dp.strftime('%d/%m/%Y')
                        c.drawString(48, y, f"- {p.get('nome')} | {p.get('telefone')} | {p.get('valor')} | {dp}")
                        y -= 14
                        if y < 80:
                            c.showPage()
                            y = h - 40
                else:
                    c.drawString(48, y, "Nenhum parceiro cadastrado.")
                    y -= 18

                c.line(40, y, w-40, y)
                y -= 24
                if y < 120:
                    c.showPage()
                    y = h - 40
            c.save()
            QMessageBox.information(self, "Sucesso", f"PDF salvo em: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar PDF: {e}")

    def _docs_to_dataframe(self, docs):
        rows = []
        for d in docs:
            descricao = d.get('descricao')
            if not descricao:
                parts = []
                if d.get('categoria'):
                    parts.append(d.get('categoria'))
                if d.get('subcategoria'):
                    parts.append(d.get('subcategoria'))
                if d.get('especificacao'):
                    parts.append(d.get('especificacao'))
                descricao = '/'.join(parts)
            row = {
                'pedido': d.get('pedido'),
                'cliente': d.get('cliente'),
                'data': d.get('data').strftime('%d/%m/%Y') if hasattr(d.get('data'), 'strftime') else d.get('data'),
                'endereco': d.get('endereco'),
                'telefone': d.get('telefone'),
                'tipo_pedido': d.get('tipo_pedido'),
                'descricao': descricao,
                'status': d.get('status'),
                'valor_tipo': d.get('valor', {}).get('tipo'),
                'valor_total': d.get('valor', {}).get('valor_total'),
                'parcelas': d.get('valor', {}).get('parcelas', 1)
            }
            rows.append(row)
        if pd:
            return pd.DataFrame(rows)
        return rows


#  Tela de Parceiros
class TelaParceiros(QWidget):
    def __init__(self):
        super().__init__()
        vbox = QVBoxLayout()
        vbox.setAlignment(Qt.AlignTop)
        self.setLayout(vbox)
        self.setStyleSheet("color: white;")

        title = QLabel("Parceiros (todos os pedidos)")
        title.setStyleSheet("font-weight: bold; font-size: 28px; color: white;")
        title.setAlignment(Qt.AlignCenter)
        vbox.addWidget(title)
        vbox.addSpacing(8)

        # filtros simples
        filt_row = QHBoxLayout()
        self.filter_text = QLineEdit()
        self.filter_text.setPlaceholderText("Filtrar por nome/cliente/pedido...")
        self.filter_text.setStyleSheet(LIGHT_BLUE_BG)
        self.filter_text.setMinimumHeight(COMMON_MIN_HEIGHT)
        btn_filter = QPushButton("Filtrar")
        btn_filter.setStyleSheet(BUTTON_STYLE)
        btn_filter.setMinimumHeight(COMMON_MIN_HEIGHT)
        btn_filter.clicked.connect(self.load_parceiros)
        filt_row.addWidget(self.filter_text)
        filt_row.addWidget(btn_filter)
        vbox.addLayout(filt_row)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(['Pedido', 'Cliente', 'Parceiro', 'Telefone', 'Valor', 'Data Pagamento', 'Descrição'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        vbox.addWidget(self.table)

        actions = QHBoxLayout()
        self.btn_refresh = QPushButton("Atualizar")
        self.btn_refresh.setStyleSheet(BUTTON_STYLE)
        self.btn_refresh.setMinimumHeight(COMMON_MIN_HEIGHT)
        self.btn_refresh.clicked.connect(self.load_parceiros)
        self.btn_export = QPushButton("Exportar CSV")
        self.btn_export.setStyleSheet(BUTTON_STYLE)
        self.btn_export.setMinimumHeight(COMMON_MIN_HEIGHT)
        self.btn_export.clicked.connect(self.export_all_csv)
        actions.addWidget(self.btn_refresh)
        actions.addWidget(self.btn_export)
        vbox.addLayout(actions)

        # carregar inicialmente
        self.load_parceiros()

    def load_parceiros(self):
        q = self.filter_text.text().strip()
        filtro = {}
        if q:
            filtro['$or'] = [
                {"parceiros.lista.nome": {"$regex": q, "$options": "i"}},
                {"parceiros.lista.telefone": {"$regex": q, "$options": "i"}},
                {"cliente": {"$regex": q, "$options": "i"}},
                {"pedido": {"$regex": q, "$options": "i"}},
                {"descricao": {"$regex": q, "$options": "i"}}
            ]
        docs = get_orders(filtro)
        rows = []
        for d in docs:
            pedido = d.get('pedido')
            cliente = d.get('cliente')
            descricao = d.get('descricao') or '/'.join(filter(None, [d.get('categoria'), d.get('subcategoria'), d.get('especificacao')]))
            parceiros = d.get('parceiros', {}).get('lista', [])
            for p in parceiros:
                dp = p.get('data_pagamento')
                dp_str = dp.strftime('%d/%m/%Y') if hasattr(dp, 'strftime') else (dp or '')
                rows.append({
                    'pedido': pedido,
                    'cliente': cliente,
                    'nome': p.get('nome'),
                    'telefone': p.get('telefone'),
                    'valor': p.get('valor'),
                    'data_pagamento': dp_str,
                    'descricao': descricao
                })
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(str(r['pedido'])))
            self.table.setItem(i, 1, QTableWidgetItem(r['cliente']))
            self.table.setItem(i, 2, QTableWidgetItem(r['nome']))
            self.table.setItem(i, 3, QTableWidgetItem(r['telefone']))
            self.table.setItem(i, 4, QTableWidgetItem(str(r['valor'])))
            self.table.setItem(i, 5, QTableWidgetItem(r['data_pagamento']))
            self.table.setItem(i, 6, QTableWidgetItem(r['descricao']))

    def export_all_csv(self):
        # pega todas as linhas mostradas e exporta
        rows = []
        for r in range(self.table.rowCount()):
            rows.append({
                'pedido': self.table.item(r, 0).text() if self.table.item(r, 0) else '',
                'cliente': self.table.item(r, 1).text() if self.table.item(r, 1) else '',
                'nome': self.table.item(r, 2).text() if self.table.item(r, 2) else '',
                'telefone': self.table.item(r, 3).text() if self.table.item(r, 3) else '',
                'valor': self.table.item(r, 4).text() if self.table.item(r, 4) else '',
                'data_pagamento': self.table.item(r, 5).text() if self.table.item(r, 5) else '',
                'descricao': self.table.item(r, 6).text() if self.table.item(r, 6) else ''
            })
        if not rows:
            QMessageBox.information(self, "Aviso", "Nenhum parceiro para exportar.")
            return
        if pd is None:
            QMessageBox.warning(self, "Dependência", "pandas não está instalado. Instale com: pip install pandas")
            return
        df = pd.DataFrame(rows)
        path, _ = QFileDialog.getSaveFileName(self, "Salvar CSV", filter="CSV (*.csv)")
        if not path:
            return
        try:
            df.to_csv(path, index=False)
            QMessageBox.information(self, "Sucesso", f"Exportado para {path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar: {e}")


#  Janela Principal
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Banco Ambiental')
        self.setMinimumSize(1200, 700)
        self.resize(1400, 900)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.central_widget.setStyleSheet("background-color: #1B3C53; color: white;")

        main_vbox = QVBoxLayout()
        self.central_widget.setLayout(main_vbox)

        top_bar = QWidget()
        top_bar_layout = QHBoxLayout()
        top_bar_layout.setContentsMargins(8, 8, 8, 8)
        top_bar_layout.setSpacing(8)
        top_bar.setLayout(top_bar_layout)
        top_bar.setFixedHeight(72)
        top_bar.setStyleSheet("background-color: #234C6A;")

        self.button_regorder = QPushButton('Registrar Pedido')
        self.button_searchorder = QPushButton('Buscar Pedido')
        self.button_parceiros = QPushButton('Parceiros')

        for btn in (self.button_regorder, self.button_searchorder, self.button_parceiros):
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
            btn.setMinimumWidth(200)
            btn.setStyleSheet(BUTTON_STYLE)
            btn.setMinimumHeight(COMMON_MIN_HEIGHT)

        top_bar_layout.addWidget(self.button_regorder)
        top_bar_layout.addWidget(self.button_searchorder)
        top_bar_layout.addWidget(self.button_parceiros)
        top_bar_layout.addStretch(1)

        main_vbox.addWidget(top_bar)

        self.content_area = QStackedWidget()
        main_vbox.addWidget(self.content_area, 1)

        self.tela_registrar = TelaRegistrarPedido()
        self.tela_buscar = TelaBuscarPedido()
        self.tela_parceiros = TelaParceiros()

        self.content_area.addWidget(self.tela_registrar)
        self.content_area.addWidget(self.tela_buscar)
        self.content_area.addWidget(self.tela_parceiros)

        self.button_regorder.clicked.connect(lambda: self.content_area.setCurrentWidget(self.tela_registrar))
        self.button_searchorder.clicked.connect(lambda: self.content_area.setCurrentWidget(self.tela_buscar))
        self.button_parceiros.clicked.connect(lambda: self.content_area.setCurrentWidget(self.tela_parceiros))

        self.content_area.setCurrentWidget(self.tela_registrar)


if __name__ == '__main__':
    window = MainWindow()
    window.show()
    app.exec()
