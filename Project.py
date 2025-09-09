from PyQt5 import QtCore, QtGui, QtWidgets
import sys, os
import mysql.connector
import networkx as nx
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook

# ------------------ Database connection ------------------
def get_db_connection():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="123456789",
        database="production_manager",
        use_pure=True
    )

# ------------------ Data model ------------------
class Stage:
    def __init__(self, sid, ops=None, deps=None, done=False):
        self.sid = sid
        self.ops = ops or []
        self.deps = deps or []
        self.done = done

class Part:
    def __init__(self, pid, order_id):
        self.pid = pid
        self.order_id = order_id
        self.stages = {}

parts = {}  # in-memory cache

# ------------------ Load from DB ------------------
def load_from_db():
    global parts
    parts.clear()
    conn = cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)

        cursor.execute("SELECT part_id, order_id FROM parts")
        for row in cursor.fetchall():
            key = (row["part_id"], row["order_id"])
            parts[key] = Part(row["part_id"], row["order_id"])

        cursor.execute("SELECT part_id, order_id, stage_id, operations, dependencies, done, operator_first, operator_last FROM stages")
        for row in cursor.fetchall():
            key = (row["part_id"], row["order_id"])
            if key not in parts:
                continue
            ops = row["operations"].split(",") if row["operations"] else []
            deps = row["dependencies"].split(",") if row["dependencies"] else []
            done = bool(row["done"])
            parts[key].stages[row["stage_id"]] = Stage(row["stage_id"], ops, deps, done)

    except mysql.connector.Error as e:
        print(f"❌ Database error while loading: {e}")
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

# ------------------ Save functions ------------------
def add_part(pid, order_id):
    if not pid or not order_id:
        return "⚠ Part ID and Order ID cannot be blank."
    key = (pid, order_id)
    if key in parts:
        return f"❌ Part {pid} with order {order_id} already exists."

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO parts (part_id, order_id) VALUES (%s, %s)", (pid, order_id))
        conn.commit()
        cursor.close()
        conn.close()
        parts[key] = Part(pid, order_id)
        return f"✅ Part {pid} (Order {order_id}) added."
    except mysql.connector.IntegrityError:
        return f"❌ Part {pid} with order {order_id} already exists in DB."
    except mysql.connector.Error as e:
        return f"❌ Database error: {e}"
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

def add_stage(pid, order_id, sid, ops, deps, operator_first="", operator_last=""):
    if not pid or not order_id or not sid:
        return "⚠ Part ID, Order ID, and Stage ID cannot be blank."

    key = (pid, order_id)
    if key not in parts:
        return f"❌ Part {pid} with order {order_id} not found."
    if sid in parts[key].stages:
        return f"❌ Stage {sid} already exists for part {pid}."

    ops_list = [o.strip() for o in ops.split(",") if o.strip()]
    deps_list = [d.strip() for d in deps.split(",") if d.strip()]

    conn = cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO stages (part_id, order_id, stage_id, operations, dependencies, done, operator_first, operator_last)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """, (pid, order_id, sid, ",".join(ops_list), ",".join(deps_list), False, operator_first, operator_last))
        conn.commit()

        parts[key].stages[sid] = Stage(sid, ops_list, deps_list, False)
        return f"✅ Stage {sid} added to part {pid}."
    except mysql.connector.Error as e:
        return f"❌ Database error: {e}"
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

def complete_stage(pid, order_id, sid):
    key = (pid, order_id)
    if key not in parts:
        return f"❌ Part {pid} (Order {order_id}) not found."
    if sid not in parts[key].stages:
        return f"❌ Stage {sid} not found in part {pid}."

    stage = parts[key].stages[sid]
    for dep in stage.deps:
        if dep not in parts[key].stages or not parts[key].stages[dep].done:
            return f"❌ Dependency {dep} not completed."

    conn = cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE stages SET done = TRUE
            WHERE part_id = %s AND order_id = %s AND stage_id = %s
        """, (pid, order_id, sid))
        conn.commit()
        stage.done = True
        return f"✅ Stage {sid} completed."
    except mysql.connector.Error as e:
        return f"❌ Database error: {e}"
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

def remove_part(pid, order_id):
    key = (pid, order_id)
    if key not in parts:
        return f"❌ Part {pid} (Order {order_id}) not found."

    conn = cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM parts WHERE part_id = %s AND order_id = %s", (pid, order_id))
        conn.commit()
        del parts[key]
        return f"✅ Part {pid} (Order {order_id}) removed."
    except mysql.connector.Error as e:
        return f"❌ Database error: {e}"
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

def remove_stage(pid, order_id, sid):
    key = (pid, order_id)
    if key not in parts or sid not in parts[key].stages:
        return f"❌ Stage {sid} not found for part {pid}."

    conn = cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            DELETE FROM stages WHERE part_id = %s AND order_id = %s AND stage_id = %s
        """, (pid, order_id, sid))
        conn.commit()
        del parts[key].stages[sid]
        return f"✅ Stage {sid} removed from part {pid}."
    except mysql.connector.Error as e:
        return f"❌ Database error: {e}"
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

def update_stage(pid, order_id, sid, ops, deps):
    key = (pid, order_id)
    if key not in parts or sid not in parts[key].stages:
        return f"❌ Stage {sid} not found in part {pid}."

    ops_list = [o.strip() for o in ops.split(",") if o.strip()]
    deps_list = [d.strip() for d in deps.split(",") if d.strip()]

    conn = cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE stages SET operations = %s, dependencies = %s
            WHERE part_id = %s AND order_id = %s AND stage_id = %s
        """, (",".join(ops_list), ",".join(deps_list), pid, order_id, sid))
        conn.commit()
        stage = parts[key].stages[sid]
        stage.ops, stage.deps = ops_list, deps_list
        return f"✅ Stage {sid} updated in part {pid}."
    except mysql.connector.Error as e:
        return f"❌ Database error: {e}"
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

# ------------------ PDF functions ------------------
def generate_part_dependency_pdf(key):
    try:
        from networkx.drawing.nx_agraph import graphviz_layout
    except ImportError:
        return "❌ Graphviz is not installed or not configured properly. Please install Graphviz and pygraphviz."

    if key not in parts:
        return f"❌ Part {key[0]} (Order {key[1]}) not found."

    part = parts[key]
    if not part.stages:
        return f"⚠ Part {key[0]} (Order {key[1]}) has no stages."

    G = nx.DiGraph()
    for stage in part.stages.values():
        G.add_node(stage.sid)
        for dep in stage.deps:
            G.add_edge(dep, stage.sid)

    pos = graphviz_layout(G, prog="dot")

    plt.figure(figsize=(6, 4))
    nx.draw(
        G, pos,
        with_labels=True,
        node_color='lightblue',
        edge_color='gray',
        node_size=2000,
        font_size=10,
        arrowsize=20
    )

    img_path = f"{key[0]}_{key[1]}_dependency_graph.png"
    plt.savefig(img_path, bbox_inches='tight')
    plt.close()

    pdf_path = f"{key[0]}_{key[1]}_dependency_graph.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    elements = [Image(img_path, width=400, height=300), Spacer(1, 12)]
    doc.build(elements)
    os.remove(img_path)
    return f"✅ Dependency graph PDF created: {pdf_path}"


def generate_completed_parts_pdf():
    pdf_filename = "completed_parts_report.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
    styles = getSampleStyleSheet()
    content = [Paragraph("Completed Parts Report", styles['Heading1']), Spacer(1, 12)]
    data = [["Part ID", "Order ID", "Completed Stages"]]
    for (pid, order_id), p in parts.items():
        if p.stages and all(s.done for s in p.stages.values()):
            data.append([pid, order_id, ", ".join(s.sid for s in p.stages.values())])
    if len(data) == 1:
        data.append(["-", "-", "-"])
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    content.append(table)
    doc.build(content)
    return f"✅ PDF generated: {pdf_filename}"


def generate_multiple_orders_report():
    pdf_filename = "multiple_orders_report.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
    styles = getSampleStyleSheet()
    content = [Paragraph("Multiple Orders Report", styles['Heading1']), Spacer(1, 12)]
    order_counts = {}
    for (pid, _order_id), _part in parts.items():
        order_counts[pid] = order_counts.get(pid, 0) + 1
    data = [["Part ID", "Order Count"]]
    for pid, count in order_counts.items():
        if count > 1:
            data.append([pid, str(count)])
    if len(data) == 1:
        data.append(["-", "-"])
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    content.append(table)
    doc.build(content)
    return f"✅ PDF generated: {pdf_filename}"

# ------------------ GUI ------------------
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)

        self.centralwidget = QtWidgets.QWidget(MainWindow)

        # ---- Main Layout ----
        mainLayout = QtWidgets.QVBoxLayout(self.centralwidget)

        # ---- Part Info ----
        partBox = QtWidgets.QGroupBox("Part Info")
        partLayout = QtWidgets.QHBoxLayout(partBox)
        partLayout.addWidget(QtWidgets.QLabel("Part ID:"))
        self.partInput = QtWidgets.QLineEdit()
        partLayout.addWidget(self.partInput)
        partLayout.addWidget(QtWidgets.QLabel("Order ID:"))
        self.orderInput = QtWidgets.QLineEdit()
        partLayout.addWidget(self.orderInput)
        mainLayout.addWidget(partBox)

        # ---- Stage Info ----
        stageBox = QtWidgets.QGroupBox("Stage Info")
        stageLayout = QtWidgets.QGridLayout(stageBox)
        stageLayout.addWidget(QtWidgets.QLabel("Stage ID:"), 0, 0)
        self.stageInput = QtWidgets.QLineEdit()
        stageLayout.addWidget(self.stageInput, 0, 1)
        stageLayout.addWidget(QtWidgets.QLabel("Operations:"), 0, 2)
        self.opsInput = QtWidgets.QLineEdit()
        stageLayout.addWidget(self.opsInput, 0, 3)
        stageLayout.addWidget(QtWidgets.QLabel("Dependencies:"), 1, 0)
        self.depsInput = QtWidgets.QLineEdit()
        stageLayout.addWidget(self.depsInput, 1, 1, 1, 3)
        mainLayout.addWidget(stageBox)

        # ---- Operator Info ----
        operatorBox = QtWidgets.QGroupBox("Operator Info")
        operatorLayout = QtWidgets.QHBoxLayout(operatorBox)
        operatorLayout.addWidget(QtWidgets.QLabel("First Name:"))
        self.firstNameInput = QtWidgets.QLineEdit()
        operatorLayout.addWidget(self.firstNameInput)
        operatorLayout.addWidget(QtWidgets.QLabel("Last Name:"))
        self.lastNameInput = QtWidgets.QLineEdit()
        operatorLayout.addWidget(self.lastNameInput)
        mainLayout.addWidget(operatorBox)

        # ---- Action Buttons ----
        buttonLayout = QtWidgets.QGridLayout()
        self.addPartBtn = QtWidgets.QPushButton("Add Part")
        self.addStageBtn = QtWidgets.QPushButton("Add Stage")
        self.completeStageBtn = QtWidgets.QPushButton("Complete Stage")
        self.listPartsBtn = QtWidgets.QPushButton("List Parts")
        buttonLayout.addWidget(self.addPartBtn, 0, 0)
        buttonLayout.addWidget(self.addStageBtn, 0, 1)
        buttonLayout.addWidget(self.completeStageBtn, 0, 2)
        buttonLayout.addWidget(self.listPartsBtn, 0, 3)
        mainLayout.addLayout(buttonLayout)

        # ---- PDF Buttons ----
        pdfLayout = QtWidgets.QHBoxLayout()
        self.depPdfBtn = QtWidgets.QPushButton("Part Dependency PDF")
        self.completedPdfBtn = QtWidgets.QPushButton("Completed Parts Report")
        self.multiPdfBtn = QtWidgets.QPushButton("Multiple Orders Report")
        pdfLayout.addWidget(self.depPdfBtn)
        pdfLayout.addWidget(self.completedPdfBtn)
        pdfLayout.addWidget(self.multiPdfBtn)
        mainLayout.addLayout(pdfLayout)

        # ---- Remove/Update Buttons ----
        crudLayout = QtWidgets.QHBoxLayout()
        self.removePartBtn = QtWidgets.QPushButton("Remove Part")
        self.removeStageBtn = QtWidgets.QPushButton("Remove Stage")
        self.updateStageBtn = QtWidgets.QPushButton("Update Stage")
        crudLayout.addWidget(self.removePartBtn)
        crudLayout.addWidget(self.removeStageBtn)
        crudLayout.addWidget(self.updateStageBtn)
        mainLayout.addLayout(crudLayout)

        # ---- Excel Report Button (hidden at first) ----
        self.excelBtn = QtWidgets.QPushButton("Generate Excel Report")
        self.excelBtn.setVisible(False)  # hidden until list is clicked
        self.excelBtn.setFixedWidth(self.excelBtn.sizeHint().width() + 20)  # make it compact

        # Apply a "bulging" style
        self.excelBtn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;   /* green */
                color: white;
                font-weight: bold;
                border-radius: 12px;
                padding: 6px 5px;
            }
            QPushButton:hover {
                background-color: #45a049;   /* darker green */
            }
            QPushButton:pressed {
                background-color: #3e8e41;   /* even darker when clicked */
            }
        """)

        mainLayout.addWidget(self.excelBtn)

        # ---- Table for displaying parts ----
        self.partsTable = QtWidgets.QTableWidget()
        self.partsTable.setColumnCount(6)
        self.partsTable.setHorizontalHeaderLabels(
            ["Part ID", "Order ID", "Stage ID", "Operations", "Dependencies", "Done"]
        )
        mainLayout.addWidget(self.partsTable)

        MainWindow.setCentralWidget(self.centralwidget)
        MainWindow.setWindowTitle("Production Manager")

# ------------------ Application ------------------
class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.addPartBtn.clicked.connect(self.handle_add_part)
        self.addStageBtn.clicked.connect(self.handle_add_stage)
        self.completeStageBtn.clicked.connect(self.handle_complete_stage)
        self.listPartsBtn.clicked.connect(self.handle_list_parts)
        self.depPdfBtn.clicked.connect(self.handle_dep_pdf)
        self.completedPdfBtn.clicked.connect(self.handle_completed_pdf)
        self.multiPdfBtn.clicked.connect(self.handle_multi_pdf)
        self.removePartBtn.clicked.connect(self.handle_remove_part)
        self.removeStageBtn.clicked.connect(self.handle_remove_stage)
        self.updateStageBtn.clicked.connect(self.handle_update_stage)

    def handle_add_part(self):
        msg = add_part(self.partInput.text().strip(), self.orderInput.text().strip())
        QtWidgets.QMessageBox.information(self, "Add Part", msg)

    def handle_add_stage(self):
        pid = self.partInput.text().strip()
        order_id = self.orderInput.text().strip()
        sid = self.stageInput.text().strip()
        ops = self.opsInput.text().strip()
        deps = self.depsInput.text().strip()
        operator_first = self.firstNameInput.text().strip()
        operator_last = self.lastNameInput.text().strip()

        msg = add_stage(pid, order_id, sid, ops, deps, operator_first, operator_last)
        QtWidgets.QMessageBox.information(self, "Add Stage", msg)

    def handle_complete_stage(self):
        msg = complete_stage(
            self.partInput.text().strip(),
            self.orderInput.text().strip(),
            self.stageInput.text().strip()
        )
        QtWidgets.QMessageBox.information(self, "Complete Stage", msg)

    def handle_list_parts(self):
    # If too many parts, export directly to Excel
        if len(parts) > 100:
            filename = generate_excel_report()
            QtWidgets.QMessageBox.warning(self, "Too Many Parts",
                f"There are more than 100 parts.\n"
                f"Data was exported to {filename} instead of showing in the table.")
            return

        # Show Excel button
        self.excelBtn.setVisible(True)
        self.excelBtn.clicked.disconnect() if self.excelBtn.receivers(self.excelBtn.clicked) > 0 else None
        self.excelBtn.clicked.connect(lambda: QtWidgets.QMessageBox.information(
            self, "Excel Report", f"✅ Excel file created: {generate_excel_report()}"))

        # Fill the table
        self.partsTable.setRowCount(0)
        row = 0
        for (pid, order_id), p in parts.items():
            for sid, stage in p.stages.items():
                self.partsTable.insertRow(row)
                self.partsTable.setItem(row, 0, QtWidgets.QTableWidgetItem(pid))
                self.partsTable.setItem(row, 1, QtWidgets.QTableWidgetItem(order_id))
                self.partsTable.setItem(row, 2, QtWidgets.QTableWidgetItem(sid))
                self.partsTable.setItem(row, 3, QtWidgets.QTableWidgetItem(", ".join(stage.ops)))
                self.partsTable.setItem(row, 4, QtWidgets.QTableWidgetItem(", ".join(stage.deps)))
                self.partsTable.setItem(row, 5, QtWidgets.QTableWidgetItem("✔" if stage.done else "✘"))
                row += 1

        if row == 0:
            QtWidgets.QMessageBox.information(self, "Parts", "⚠ No parts available.")

    def handle_dep_pdf(self):
        msg = generate_part_dependency_pdf((self.partInput.text().strip(), self.orderInput.text().strip()))
        QtWidgets.QMessageBox.information(self, "PDF", msg)

    def handle_completed_pdf(self):
        msg = generate_completed_parts_pdf()
        QtWidgets.QMessageBox.information(self, "PDF", msg)

    def handle_multi_pdf(self):
        msg = generate_multiple_orders_report()
        QtWidgets.QMessageBox.information(self, "PDF", msg)

    def handle_remove_part(self):
        pid, order_id = self.partInput.text().strip(), self.orderInput.text().strip()
        reply = QtWidgets.QMessageBox.question(self, "Confirm",
            f"Delete part {pid} (Order {order_id}) and all its stages?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            msg = remove_part(pid, order_id)
            QtWidgets.QMessageBox.information(self, "Remove Part", msg)

    def handle_remove_stage(self):
        pid, order_id, sid = self.partInput.text().strip(), self.orderInput.text().strip(), self.stageInput.text().strip()
        reply = QtWidgets.QMessageBox.question(self, "Confirm",
            f"Delete stage {sid} from part {pid} (Order {order_id})?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.Yes:
            msg = remove_stage(pid, order_id, sid)
            QtWidgets.QMessageBox.information(self, "Remove Stage", msg)

    def handle_update_stage(self):
        msg = update_stage(
            self.partInput.text().strip(),
            self.orderInput.text().strip(),
            self.stageInput.text().strip(),
            self.opsInput.text().strip(),
            self.depsInput.text().strip()
        )
        QtWidgets.QMessageBox.information(self, "Update Stage", msg)

# -------------- Excel generator function ---------------
def generate_excel_report(filename="parts_report.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Parts Report"

    headers = ["Part ID", "Order ID", "Stage ID", "Operations", "Dependencies", "Done"]
    ws.append(headers)

    for (pid, order_id), p in parts.items():
        for sid, stage in p.stages.items():
            ws.append([
                pid,
                order_id,
                sid,
                ", ".join(stage.ops),
                ", ".join(stage.deps),
                "✔" if stage.done else "✘"
            ])

    wb.save(filename)
    return filename

# ------------------ Run ------------------
if __name__ == "__main__":
    load_from_db()
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())