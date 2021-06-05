#!/usr/bin/python3
#_*_ coding: utf-8 _*_

from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import *

from bs4 import BeautifulSoup, SoupStrainer # Permite encontrar los links en el archivo de forma facil
import requests    # Solicitar las paginas web
import html2text   # Transformar el HTML a texto plano.
import xlsxwriter  # Crear el archivo para Excel

# Includidas por defecto en python 
import sys         # sistema 
import re          # Trabajar con cadenas de texto
import csv         # Crear el csv



only_parse_species      = SoupStrainer('div',attrs={'class':"speciesTitle"})
only_parse_family_links = SoupStrainer('a',attrs={'title':'Classic view'})

text_maker = html2text.HTML2Text()
text_maker.ignore_links        = True
text_maker.bypass_tables       = False
text_maker.ignore_emphasis     = True 
text_maker.images_to_alt       = True 
text_maker.ignore_anchors      = True  
text_maker.single_line_break   = True 
text_maker.use_automatic_links = True 

# Puedesdecir "¿Por qué no mejor hacer una clase para esto? pero me daba flojera...
columnas_descriptivas = ['Specie','Autor','Family','Genus','Described Sex','Distribution','LSID','WSC URL']

class RequestEstadisticas(QtCore.QThread):
    signal = QtCore.pyqtSignal(object)
    signal_p = QtCore.pyqtSignal(object)
    
    def __init__(self, parent=None):
        super(RequestEstadisticas, self).__init__(parent)
        self._stopped = True
        self._mutex = QtCore.QMutex()

    def stop(self):
        self._mutex.lock()
        self._stopped = True
        self._mutex.unlock()

    def run(self):
        self._stopped = False
        try:
            estadisticas = requests.get('https://wsc.nmbe.ch/statistics/')
        except:
            self.signal.emit(' <b>Error:</b> <i>No se puede conectar</i>')
        else:
            self.signal_p.emit('Procesando datos obtenidos')
            soup = BeautifulSoup(estadisticas.text, 'html.parser')
            table_rows = soup.find_all('tr')
            last_row =  str(table_rows[len(table_rows)-1])
            items = text_maker.handle(last_row).replace(' ','').split("|")
            numero_familias = len(table_rows)-2
            data = [int(items[3]),numero_familias,int(items[2])]
            self.signal.emit(data)

class RequestFamiliasValidas(QtCore.QThread):
    signal = QtCore.pyqtSignal(object)
    signal_p = QtCore.pyqtSignal(object)
    
    def __init__(self, parent=None):
        super(RequestFamiliasValidas, self).__init__(parent)
        self._stopped = True
        self._mutex = QtCore.QMutex()

    def stop(self):
        self._mutex.lock()
        self._stopped = True
        self._mutex.unlock()

    def run(self):
        self._stopped = False
        try:
            page_families = requests.get('https://wsc.nmbe.ch/families')
        except:
            self.signal.emit(' <b>Error:</b> <i>No se puede conectar</i>')
        else:
            self.signal_p.emit('Procesando datos obtenidos...')
            wsc_html_soup = BeautifulSoup(page_families.content, 'html.parser', parse_only=only_parse_family_links)
            family_link_list = wsc_html_soup.find_all('a',attrs={'title':'Classic view'})
            self.signal.emit(family_link_list)

class ProcesarFamilia(QtCore.QThread):
    signal = QtCore.pyqtSignal(object)
    signal_m = QtCore.pyqtSignal(object)
    link = ''
    family_name = '' 
    
    def __init__(self, parent=None):
        super(ProcesarFamilia, self).__init__(parent)
        self._stopped = True
        self._mutex = QtCore.QMutex()
    
    def setThreadInfo(self,link, name):
        self.link = link
        self.family_name = name

    def stop(self):
        self._mutex.lock()
        self._stopped = True
        self._mutex.unlock()

    def run(self):
        self._stopped = False
        
        text = ' Solicitando información de <i>'+self.family_name+'</i><br>'
        text = text+'desde <b><i>'+self.link+'</b></i>'
        self.signal_m.emit([1,text])
        try:
            page_families = requests.get(self.link).content
        except:
            self.signal_m.emit([0,' <b>Error:</b> <i>No se puede conectar</i>'])
        else:
            text = ' Ejecutando analizador sintáctico en los datos de <b><i>'+self.family_name+'</i></b>'
            
            self.signal_m.emit([1,text])
            
            family_html_soup = BeautifulSoup(page_families, 'html.parser',  parse_only=only_parse_species)
            # saquemos todas las especies en esta familia... class="speciesTitle"
            species_list = family_html_soup.find_all('div',attrs={'class':"speciesTitle"})
            
            text = ' Procesando las especies de <b><i>'+self.family_name+'</i></b>'
            self.signal_m.emit([1,text])
            specie_num = 0
            for specie in species_list:
                specie_num += 1
                self.signal_m.emit([2])
                
                fragmento = str(specie)
                text = text_maker.handle(fragmento).lstrip()
                text = text.replace('\n','').replace('[',' [')
                text = text.replace('| ? |','| Unknow |')

                # expresiones regulares... poderosas y odiosas expresiones regulares
                described = re.search('\|(.*)\|', text ).group(1)
                full_d = '\|'+described+'\|'
                distribution = re.search(full_d+'(.*)\[', text).group(1).lstrip().replace('\n','')
                
                sd_distribution = distribution.lower()
                
                IS_ON_MEXICO = re.search("m\s*[eé]\s*[xj]\s*i\s*c\s*o", sd_distribution)
                if not IS_ON_MEXICO :
                    #coverting all possible space typos...
                    IS_ON_MEXICO = re.search("n\s*o\s*r\s*t\s*h\s*a\s*m\s*e\s*r\s*i\s*c\s*a", sd_distribution)
                    if not IS_ON_MEXICO :
                        IS_ON_MEXICO = re.search("usa\s*to", sd_distribution)
                        if not IS_ON_MEXICO:
                            IS_ON_MEXICO = (
                                re.search("canada\s*to", sd_distribution) 
                                and not re.search("canada\s*to\s*usa", sd_distribution)
                                )
                
                if IS_ON_MEXICO:
                    # hacer el resto del procesamiento para escribir los datos 
                    
                    specie_name = specie.a.text
                    specie_url = specie.a['href'].replace( '/'+specie_name.replace(' ','_') ,'')
                    specie_url = 'https://wsc.nmbe.ch'+specie_url
                    genus = specie_name.split()[0] 
                
                    described = described.replace("m", "M♂").replace("f", " F♀").lstrip()
                    
                    autor = re.search(specie_name+'(.*)'+full_d, text ).group(1)
                    autor = autor.replace('*','').lstrip()
                    
                    start = text.find("[") + len("[")
                    end = text.find("]")
                    LSID = '['+text[start:end]+']'
                    
                
                    lista_campos = [specie_name,autor,self.family_name,genus,described,distribution,LSID,specie_url]
                    
                    self.signal_m.emit([3,lista_campos])
                else:
                    pass
                    #print(specie_name+' is not on Mexico or North America')
                
                
            self.signal_m.emit([4])
            
            

class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        QMainWindow.__init__(self, *args, **kwargs)
        self.tabla_de_aracnidos = None
        self.no_starting_info_label = None
        self.topFrame = None
        #self.main_info_label = None
        self.panel_dividido = None

        self.is_tabla_aracnidos_iniciada = False
        self.familia_list_raw_data = None
        self.familia_procesed_id = 0
        
        self.lista_de_aracnidos_mexicanos = []
        self.datos_completos = False

        self.especies_analizadas = 0
        self.especies_en_mexico = 0
        self.numero_total_de_especies = 0
        self.numero_total_de_familias = 0
        self.numero_total_de_generos = 0
        self.text_main_info_label = "Numero total de especies: <b>{}</b><br>Numero total de familias: <b>{}</b><br>Numero total de generos: <b>{}</b>"
        self.text_especies =  "Especies comparadas: <b>{}</b><br>Especies en México: <b>{}</b><br>"
        
        self.hiloProcesarFamilia = None

    def build(self):
        #"application-exit" 
        exitAct = QAction(QtGui.QIcon.fromTheme("window-close"), 'Salir', self)
        exitAct.setShortcut('Ctrl+Q')
        exitAct.triggered.connect(qApp.quit)
        
        saveAsExcelAct = QAction(QtGui.QIcon.fromTheme("x-office-spreadsheet"), 'Guardar como Excel', self)
        saveAsExcelAct.triggered.connect(self.saveExcelDialog)
        
        saveAsCSVAct = QAction(QtGui.QIcon.fromTheme("text-x-generic"), 'Guardar como CSV', self)
        saveAsCSVAct.triggered.connect(self.saveCSVDialog)
        
        self.toolbar = self.addToolBar('Herramientas')
        self.toolbar.setFloatable(False)
        self.toolbar.setMovable(False) 
        self.toolbar.setStyleSheet("QToolBar{spacing:0px;}");
        self.toolbar.addAction(saveAsExcelAct)
        self.toolbar.addAction(saveAsCSVAct)
        self.toolbar.addSeparator()
        
        placeholder = QWidget(self) 
        placeholder.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.toolbar.insertWidget(QAction('',self),placeholder)
        self.toolbar.addAction(exitAct)
    
        layout = QGridLayout()
        layoutTopFrame = QGridLayout()
        

        self.topFrame = QFrame()

        self.request_info_label = QLabel('No data, no data...')
        #self.request_info_label.setFont(QtGui.QFont("Arial", 10))
        self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;")
        #self.topFrame.setSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        #self.topFrame.setAlignment(QtCore.Qt.AlignCenter)
        text = self.text_main_info_label.format(
            self.numero_total_de_especies,
            self.numero_total_de_familias,
            self.numero_total_de_generos)
        self.main_info_label = QLabel(text)
        self.main_info_label.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Preferred)

        
        text = self.text_especies.format(
            self.especies_analizadas,
            self.especies_en_mexico)
        self.species_info_label = QLabel(text)
        self.species_info_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        
        #self.main_info_label.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Black))
        
        
        self.iniciar_button = QPushButton('Buscar arañas en México')
        self.iniciar_button.setEnabled(False)
        self.iniciar_button.clicked.connect(self.startRequestFamilias)
        
        self.progress = QProgressBar(self)
        self.progress.hide()
        
        layoutTopFrame.addWidget(self.main_info_label, 0,0, 2,2)
        layoutTopFrame.addWidget(self.species_info_label, 0,2, 2,2)
        layoutTopFrame.addWidget(self.iniciar_button,3,4, 1,2)

        layoutTopFrame.addWidget(self.request_info_label,3,0,1,4)
        layoutTopFrame.addWidget(self.progress,4,0, 1,6)
        
        self.topFrame.setLayout(layoutTopFrame)

        # Bottom frame...
        self.no_starting_info_label = QLabel('Aun no hay datos de especies en México')
        self.no_starting_info_label.setToolTip('Cuando comienzen a recogerse\ndatos de especies en México\napareceran aquí.')
        self.no_starting_info_label.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Expanding)
        self.no_starting_info_label.setAlignment(QtCore.Qt.AlignCenter)
        self.no_starting_info_label.setStyleSheet("background-color: grey;")
        
        self.panel_dividido = QSplitter(QtCore.Qt.Vertical)
        self.panel_dividido.addWidget(self.topFrame)
        self.panel_dividido.addWidget(self.no_starting_info_label)
        self.panel_dividido.setSizes([50,450])

        layout.addWidget(self.panel_dividido)

        self.requestEstadisticas = RequestEstadisticas()
        self.requestEstadisticas.started.connect(self.startRequestWSCEstadisticas)
        #self.requestEstadisticas.finished.connect(self.worker_finished_callback)
        self.requestEstadisticas.signal.connect(self.sendResultRequestWSCEstadisticas)
        self.requestEstadisticas.signal_p.connect(self.procesingRequestWSCEstadisticas)
        self.requestEstadisticas.start()
        
        self.requestFamilias = RequestFamiliasValidas()
        self.requestFamilias.signal.connect(self.sendResultRequestFamilias)
        self.requestFamilias.signal_p.connect(self.procesingFamilias)
        
        self.hiloProcesarFamilia = ProcesarFamilia()
        self.hiloProcesarFamilia.signal_m.connect(self.messagesProcesarFamilia)
        
        
        self.central_widget = QWidget()               # define central widget
        self.setCentralWidget(self.central_widget)  
        #self.setLayout(layout)
        self.centralWidget().setLayout(layout) 
        self.show()
        self.setWindowTitle('Arañas de México en el WSC')
        self.resize(750,500)
    
    
    
    def startRequestWSCEstadisticas(self):
        self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
        self.request_info_label.setText('Solicitando información de <b><i>https://wsc.nmbe.ch/statistics/</b></i>')
    
    def procesingRequestWSCEstadisticas(self, data):
        self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
        self.request_info_label.setText(data)
    
    def sendResultRequestWSCEstadisticas(self, data):
        if isinstance(data, str):
            #a terrible error just happend...
            self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New; background-color: yellow;") 
            self.request_info_label.setText(data)
            #Intentamos reconectarnos dentro de los proximos 2.5 segundos
            timer = QtCore.QTimer(self)
            timer.timeout.connect(self.requestEstadisticas.start)
            timer.start(2500)
        else:
            self.iniciar_button.setEnabled(True)
            self.numero_total_de_especies = data[0]
            self.numero_total_de_familias = data[1]
            self.numero_total_de_generos =  data[2]
            text = self.text_main_info_label.format(
                self.numero_total_de_especies,
                self.numero_total_de_familias,
                self.numero_total_de_generos)
            self.progress.setMaximum(self.numero_total_de_especies)
            self.main_info_label.setText(text)
            self.request_info_label.hide() #setText('Sin acciones pendientes.')
            
            
    def startRequestFamilias(self):
        self.iniciar_button.setEnabled(False)
        self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
        self.request_info_label.setText('Solicitando información de <b><i>https://wsc.nmbe.ch/families</b></i>')
        self.request_info_label.show()
        self.requestFamilias.start()
        
    def procesingFamilias(self, data):
        self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
        self.request_info_label.setText(data)
    
    def sendResultRequestFamilias(self, data):
        if isinstance(data, str):
            #a terrible error just happend...
            self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New; background-color: yellow;") 
            self.request_info_label.setText(data)
            #Intentamos reconectarnos dentro de los proximos 2.5 segundos
            timer = QtCore.QTimer(self)
            timer.timeout.connect(self.startRequestFamilias)
            timer.start(2500)
        else:
            self.familia_list_raw_data = data
            self.familia_procesed_id = 0
            self.startDownloadFamily()
            self.progress.show()
            #self.request_info_label.hide() #setText('Sin acciones pendientes.')
    
    def startDownloadFamily(self):
        family_info = self.familia_list_raw_data[self.familia_procesed_id]
        family_name = family_info['href'].split('/')[3]
        link_url = ('https://wsc.nmbe.ch'+family_info['href'])
        
        self.hiloProcesarFamilia.setThreadInfo(link_url,family_name) 
        
        self.hiloProcesarFamilia.start()
    
    def messagesProcesarFamilia(self, data):
        if data[0] == 0:
            self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New; background-color: yellow;") 
            self.request_info_label.setText(data[1])
            timer = QtCore.QTimer(self)
            timer.timeout.connect(self.startDownloadFamily)
            timer.start(2500)
        elif data[0] == 1:
            self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
            self.request_info_label.setText(data[1])
        elif data[0] == 2 or data[0] == 3:
            if data[0] == 2 :
                self.especies_analizadas += 1
                self.progress.setValue(self.especies_analizadas)
            if data[0] == 3 :
                self.especies_en_mexico += 1
                self.iniciarTablaAracnidos()
                self.agregarNuevaFilaATablaAracnidos(data[1])
                #guardar los datos de la araña
            text = self.text_especies.format(
                    self.especies_analizadas,
                    self.especies_en_mexico)
            self.species_info_label.setText(text)
        elif data[0] == 4: 
            self.revisarFamiliasPendientes()
            
    def revisarFamiliasPendientes(self):
        if self.familia_procesed_id+1 < len(self.familia_list_raw_data):
            self.familia_procesed_id += 1
            self.startDownloadFamily()
        else:
            self.request_info_label.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
            self.request_info_label.hide()
            #self.progress.hide()
            self.iniciar_button.setEnabled(True)
            self.datos_completos = True

    def iniciarTablaAracnidos(self):
        #quitar el no info label del panel inferior...
        #self.panel_dividido.removeWidget(self.no_starting_info_label)
        #self.no_starting_info_label.deleteLater()
        if not self.is_tabla_aracnidos_iniciada :
            
            self.lista_de_aracnidos_mexicanos = []
            self.datos_completos = False
            
            self.no_starting_info_label.setParent(None)
            self.no_starting_info_label = None
            
            self.tabla_de_aracnidos = QTableWidget()
            self.tabla_de_aracnidos.setColumnCount(len(columnas_descriptivas))
            self.tabla_de_aracnidos.setHorizontalHeaderLabels(columnas_descriptivas)
            self.panel_dividido.addWidget(self.tabla_de_aracnidos)
            self.panel_dividido.setSizes([100,400])
            self.is_tabla_aracnidos_iniciada = True

    def agregarNuevaFilaATablaAracnidos(self, datos_aracnido):
        self.lista_de_aracnidos_mexicanos.append(datos_aracnido)
        
        row = self.tabla_de_aracnidos.rowCount()
        self.tabla_de_aracnidos.setRowCount(row+1)
        for col in range(0,len(datos_aracnido)):
            item = QTableWidgetItem(datos_aracnido[col])
            #item.setStyleSheet(" font-size: 12px; font-family: Courier New;") 
            item.setFlags(QtCore.Qt.ItemIsEnabled)
            self.tabla_de_aracnidos.setItem( row, col, item  )
            header = self.tabla_de_aracnidos.horizontalHeader() 
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)

    def saveExcelDialog(self):
        #options |= QFileDialog.DontUseNativeDialog
        if not self.datos_completos :
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('¡No hay datos que guardar!')
            msg.setWindowTitle("Error")
            msg.exec_()
        else:
            options = QFileDialog.Options()
            fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","AracnidosMexico","Archivo Microsoft Excel 2007 (*.xlsx)", options=options)
            if fileName:
                name = fileName.lstrip()
                if name[-5:len(name)] != '.xlsx':
                    name = name+'.xlsx'
                print(name)
                workbook = xlsxwriter.Workbook(name, {'constant_memory': True})
                worksheet = workbook.add_worksheet()
                
                fila = 0
                i = 0
                while i < len(columnas_descriptivas):
                    worksheet.write(fila, i, columnas_descriptivas[i])
                    i=i+1
                fila = 1
                
                for aracnido in self.lista_de_aracnidos_mexicanos:
                    i=0
                    while i < 8:
                        worksheet.write(fila,  i, aracnido[i])
                        i=i+1
                    fila += 1
               
                workbook.close()
    
    def saveCSVDialog(self):
        #options |= QFileDialog.DontUseNativeDialog
        if not self.datos_completos :
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('¡No hay datos que guardar!')
            msg.setWindowTitle("Error")
            msg.exec_()
        else:
            options = QFileDialog.Options()
            fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","AracnidosMexico","Archivo CSV (*.csv)", options=options)
            if fileName:
                name = fileName.lstrip()
                if name[-4:len(name)] != '.csv':
                    name = name+'.csv'
                print(name)
                csv_file  = open(name, 'w', encoding='utf-8')
                csv_on_mexico = csv.writer(csv_file, dialect='excel')

                csv_on_mexico.writerow(columnas_descriptivas)
                
                for aracnido in self.lista_de_aracnidos_mexicanos:
                    csv_on_mexico.writerow(aracnido)
               
                csv_file.close()
                


if __name__ == '__main__':
    App = QApplication(sys.argv)
    window = MainWindow()
    window.build()
    if (sys.flags.interactive != 1) or not hasattr(QtCore, 'PYQT_VERSION'):
        QApplication.instance().exec_()

