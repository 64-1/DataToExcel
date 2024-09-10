import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QFileDialog, QMessageBox
from main_script import main  # Ensure your script is named `main_script.py` and the main functionality is encapsulated in a `main()` function.

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('File Processing Tool')
        self.setGeometry(300, 300, 350, 200)

        layout = QVBoxLayout()

        self.label = QLabel('Select a directory:', self)
        layout.addWidget(self.label)

        self.btnSelectFolder = QPushButton('Select Folder', self)
        self.btnSelectFolder.clicked.connect(self.openFolderDialog)
        layout.addWidget(self.btnSelectFolder)

        self.btnRun = QPushButton('Run Processing', self)
        self.btnRun.clicked.connect(self.runProcessing)
        layout.addWidget(self.btnRun)

        self.setLayout(layout)

    def openFolderDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        directory = QFileDialog.getExistingDirectory(self, 'Select a Folder', options=options)
        if directory:
            self.label.setText(f"Selected Directory: {directory}")
            # Set the directory in your script here, modifying the script to accept an external directory.
            main_script.set_directory(directory)  # You would need to adjust your script to handle this.

    def runProcessing(self):
        try:
            main()  # Call the main function of your script.
            QMessageBox.information(self, 'Success', 'Processing completed successfully!')
        except Exception as e:
            QMessageBox.warning(self, 'Error', f'An error occurred: {e}')

def main_gui():
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main_gui()
