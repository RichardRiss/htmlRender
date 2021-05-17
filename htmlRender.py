#!/usr/bin/python3

import os
from stat import *
import sys
import time
import re
from watchdog.observers import Observer
from watchdog.events import RegexMatchingEventHandler
#Apache 2.0 Lizenz (Apache Lizenz Info muss beiliegen)
#import dominate
#from dominate.tags import *
#GPL V3.0 Lizenz
import pandas as pd
import numpy as np
#BSD Lizenz
import yaml
#MIT Lizenz




#Create Class for ini Handling
class yamlHandler:
    # mit self. können Variablen für gesamte class gespeichert, weitergegeben werden
    def __init__(self):
        # os.path.realpath - gibt den dirketen Pfad ohne Verknüpfungen der angegeben Datei aus
        # os.path.split - Teilt den Pfad am letzten "\" in Head [0] und Tail [1]
        # __file__ - gibt den Pfad des aktuellen Skriptes aus
        # --> Erzeugt den Pfad der yaml-Datei aus dem Pfad des aktuellen Skripts
        self.src_path=os.path.split((os.path.realpath(__file__)))[0]+ "\hmtlRender.yaml"

    def prim_setup(self):
        # Prüfen ob die yaml-Datei im ermitteltem Pfad existiert
        # wenn ja --> ok
        # wenn nein --> Datei im aktuellem Pfad erstellen
        try:
            if os.path.exists(self.src_path):
                if not getattr(sys, 'frozen', False):
                    os.remove(self.src_path)
                    self.setup()
            else:
                self.setup()

            if os.path.exists(yamlHandler().read('plcPath') + "rxcorr.plc"):
                plcEventHandler().process(yamlHandler().read('plcPath') + "rxcorr.plc")
                #os.remove(yamlHandler().read('plcPath') + "rxcorr.plc" )
            if os.path.exists(yamlHandler().read('plcPath') + "applcorr.plc"):
                plcEventHandler().process(yamlHandler().read('plcPath') + "applcorr.plc")
                #os.remove(yamlHandler().read('plcPath') + "applcorr.plc")


        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))

    def setup(self):
        # erzeugt die yaml-Datei, falls sie noch nicht existiert
        self.data={
            'HTML_Path': ['C:\\ProgramData\\Siemens\\MotionControl\\addon\\sinumerik\\hmi\\cfg\\MW_Site.html',
                          'C:\\ProgramData\\Siemens\\MotionControl\\addon\\sinumerik\\hmi\\cfg\\TW_Site.html',
                          'C:\\ProgramData\\Siemens\\MotionControl\\addon\\sinumerik\\hmi\\cfg\\Appl_Site.html'],
            'plcPath':'C:\\ProgramData\\Siemens\\MotionControl\\user\\sinumerik\\hmi\\appl\\',
            'Anz_Zeilen':100,
            'File_End':'.plc',
            'MemFile': [os.path.split((os.path.realpath(__file__)))[0] + '\\mem_mw.csv',
                        os.path.split((os.path.realpath(__file__)))[0] + '\\mem_tw.csv',
                        os.path.split((os.path.realpath(__file__)))[0] + '\\mem_appl.csv'],
            'MaxLines':100,
            'PSS_Status':os.path.split((os.path.realpath(__file__)))[0] + '\PSS_Status.txt',
            'MW-Header': ["SerienNummer", "PSS-Teilestatus", "Datum", "Uhrzeit", "Traveltime", "MwStatus",
                          "Korrektur erforderlich", "Vollkorr.", "verwerfen",
                          "Mw1", "Mw1Status", "Korr.1", "Mw2", "Mw2Status", "Korr.2", "Mw3", "Mw3Status", "Korr.3",
                          "Mw4", "Mw4Status", "Korr.4", "Mw5", "Mw5Status", "Korr.5", "Mw6", "Mw6Status", "Korr.6",
                          "Mw7", "Mw7Status", "Korr.7", "Mw8", "Mw8Status", "Korr.8", "Mw9", "Mw9Status", "Korr.9",
                          "Mw10", "Mw10Status", "Korr.10", "Mw11", "Mw11Status", "Korr.11", "Mw12", "Mw12Status", "Korr.12",
                          "Mw13", "Mw13Status", "Korr.13", "Mw14", "Mw14Status", "Korr.14", "Mw15", "Mw15Status", "Korr.15",
                          "Mw16", "Mw16Status", "Korr.16", "Mw17", "Mw17Status", "Korr.17", "Mw18", "Mw18Status", "Korr.18",
                          "Mw19", "Mw19Status", "Korr.19", "Mw20", "Mw20Status", "Korr.20", "Mw21", "Mw21Status", "Korr.21",
                          "Mw22", "Mw22Status", "Korr.22", "Mw23", "Mw23Status", "Korr.23", "Mw24", "Mw24Status", "Korr.24",
                          "Mw25", "Mw25Status", "Korr.25"],
            'TW-Header': ["SerienNummer", "PSS-Teilestatus", "Datum", "Uhrzeit", "Traveltime", "MwStatus",
                          "Korrektur erforderlich", "Vollkorr.", "verwerfen",
                          "Tw1", "Tw1Status", "Korr.1", "Tw2", "Tw2Status", "Korr.2", "Tw3", "Tw3Status", "Korr.3",
                          "Tw4", "Tw4Status", "Korr.4", "Tw5", "Tw5Status", "Korr.5", "Tw6", "Tw6Status", "Korr.6",
                          "Tw7", "Tw7Status", "Korr.7", "Tw8", "Tw8Status", "Korr.8", "Tw9", "Tw9Status", "Korr.9",
                          "Tw10", "Tw10Status", "Korr.10", "Tw11", "Tw11Status", "Korr.11", "Tw12", "Tw12Status", "Korr.12",
                          "Tw13", "Tw13Status", "Korr.13", "Tw14", "Tw14Status", "Korr.14", "Tw15", "Tw15Status", "Korr.15",
                          "Tw16", "Tw16Status", "Korr.16", "Tw17", "Tw17Status", "Korr.17", "Tw18", "Tw18Status", "Korr.18",
                          "Tw19", "Tw19Status", "Korr.19", "Tw20", "Tw20Status", "Korr.20", "Tw21", "Tw21Status", "Korr.21",
                          "Tw22", "Tw22Status", "Korr.22", "Tw23", "Tw23Status", "Korr.23", "Tw24", "Tw24Status", "Korr.24",
                          "Tw25", "Tw25Status", "Korr.25"],
            'Status': ["ungueltiger Status (Index 0)" , "i.O.", "Nacharbeit", "Ausschuss", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve"
                , "Reserve", "Reserve", "Verworfen, i.O", "Verworfen, Nacharbeit", "Verworfen, Ausschuss", "Reserve",
                                       "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve"
                , "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve",
                                       "Reserve", "Reserve", "Reserve", "Reserve","Reserve","Reserve",
                                       "Reserve","Reserve","Ungueltig, von Messstation","Verworfen, neue Messung","Verworfen, von Messstation",
                                       "Verworfen, Korrektursperre", "Verworfen, ueberholt von HighPrio-Teil / zu alt",
                                       "Verworfen, Teil nicht in Historie"],
            'PSS_Teilestatus': ["i.O.","n.i.O.","Ausschuss","Feinmessraum","Gesperrt","Materialfehler","Nacharbeit",
                                "Stichprobe","Reserve", "Reserve", "Reserve", "Reserve","Bewertung nicht abgeschlossen",
                                "Versuch fertig","Werkzeugbruch","Fertig bearbeitet","Technologieversuch","Reserve", "Reserve",
                                "ErPr.oH","ErPr.mH","ErPr.nStil","ErPr.Tw.oH","ErPr.Tw.mH","ErPr.oH","ErPr.mH",
                                "ErPr.Wb.oH","ErPr.Wb.mH","ErPr. In Maschine","man.MassKor","ErPr.Wst.n.Mass",
                                "Letztes Wstck vor WZ-Tausch","Reserve", "Reserve", "Reserve","Prozess Aenderung","Kalibrierung",
                                "Stich.KMG","Stich.PP","MFU","PMFU","Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve","B.unvollst","Daten sind ungueltig",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve", "Reserve",
                                "Reserve", "Reserve", "Reserve", "Reserve",]
        }
        with open(self.src_path,'w',encoding='utf8') as f:  # Dokumente im aktuellem Pfad erstellen
            yaml.dump(self.data,f)                          # Daten in die yaml-Datei schreiben
        f.close()                                           # Datei schließen



    def read(self,key):                             # ...für key wird nichts übergeben
        self.key=key
        # yaml.safe_load - sicheres Laden, wenn aus nicht vertrauenswürdiger Quelle. Dann werden nur Integers oder
        # Listen erstellt. Dateien können als "sicher" markiert werden
        with open(self.src_path, 'r') as f:
            self.ret_val=yaml.safe_load(f)
            if self.ret_val is None:
                return False
            else:
                try:
                    return (self.ret_val[key])
                except KeyError:
                    print("Broken Key in Yaml File. Delete " + self.src_path + " and retry.")
                    exit()

    def add(self,key,value):
        assert key                              # Überprüft ob eine Bedingung zutrifft und gibt dann true/false zurück
                                                # hier: ist key vorhanden ?
        self.key=key
        self.value=value
        with open(self.src_path,'r') as f:
            yaml_cont=yaml.safe_load(f) or {}   # Inhalt der yaml-Datein in yaml_cont schreiben
        yaml_cont[self.key]=self.value          # aktuelle Werte in yaml_cont hinzufügen
        with open(self.src_path,'w') as f:
            yaml.dump(yaml_cont,f)              # yaml-Datei neu schreiben, jetzt mit aktuellen Werten

#Create Observer Class with own Event Class
class plcWatcher:
    def __init__(self,src_path):
        # initialisieren von Variablen die global (auch außerhalb der Class!) verwendet werden soll
        self.__src_path=src_path
        self.__event_handler=plcEventHandler()
        self.__event_observer=Observer()


    def run(self):
        self.start()                # Start der Watchdog-Funktion
        try:
            while True:             # warte immer wieder 1 sek
                time.sleep(1)
        except KeyboardInterrupt:   # Abbruch über Tastaureingabe
            self.stop()             # Stopp der Watchdog-Funktion

    def start(self):
        self.__schedule()
        self.__event_observer.start()


    def __schedule(self):
        print("Watchdog active")
        print("Looking at: " + self.__src_path + yamlHandler().read('File_End'))
        self.__event_observer.schedule(
            self.__event_handler,
            self.__src_path,
            recursive=False         # achtet nicht auf rekursive Verwänderung?
        )

    def stop(self):
        self.__event_observer.stop()
        self.__event_observer.join()

#Create Event Class
class plcEventHandler(RegexMatchingEventHandler):

    def __init__(self):
        # Class ruft Unterclass auf (plcEventHandler(RegexMatchingEventHandler):)
        # um Unterklasse von Oberklasse "erben" zu lassen : super().
        super().__init__(regexes=[r".*\.plc$"])
        #super().__init__(ignore_patterns="")
        #super().__init__(ignore_directories=False)
        #super().__init__(case_sensitive=True)

    def on_created(self,event):
        #Eventuell für Zugriffsfehler beim Erzeugen der Datei
        self.path=event.src_path
        time.sleep(1)
        acc=False
        timeout=time.time()+3 #3s timeout für Warteschleife
        # while not acc and (time.time() < timeout):
        #     try:
        #         os.rename(self.path,self.path)
        #         acc=True
        #     except:
        #         acc=False
        #         continue
        #file_size = -1
        #while file_size != os.path.getsize(event.src_path):
        #   file_size = os.path.getsize(event.src_path)
        self.process(self.path)


    def process(self,path):
        self.path=path
        Eingangs_Block = []

        #Try to read the new measurement file
        try:
            with open(self.path,'r') as f:
                for line in f:
                    Eingangs_Block.append([n for n in line.strip()])
            #Delete Meas. File
            os.remove(self.path)
        except:
        #TODO:Better error handling -> Log File Function
            print("Error in Processing")
            print(sys.exc_info()[1])
            print("Unix Permission for file: " + oct(os.stat(self.path)[ST_MODE])[-3:])

        file_indicator=os.path.basename(self.path)

        #Decode Data and write Memory File
        try:
            #Memory File for Received Correction
            if file_indicator == "rejected.plc":
                self.mem_reject=self.manipulate_mem_file(Eingangs_Block[0])
            elif file_indicator == "rxcorr.plc":
                mw_data_dec, tw_data_dec = self.Decodieren(Eingangs_Block[0], len(Eingangs_Block[0]))
                if len(mw_data_dec) > 0:
                    self.mem_mw = self.Memory_File(mw_data_dec,0)
                if len(tw_data_dec) > 0:
                    self.mem_tw = self.Memory_File(tw_data_dec,1)
            #Memory File for Applied Correction
            elif file_indicator=="applcorr.plc":
                mw_data_dec, tw_data_dec = self.Decodieren(Eingangs_Block[0], len(Eingangs_Block[0]))
                if len(mw_data_dec) > 0 and mw_data_dec[7]=="ja":
                    self.mem_appl = self.Memory_File(mw_data_dec, 2)
                elif len(tw_data_dec) > 0 and mw_data_dec[7]=="nein":
                    self.mem_appl = self.Memory_File(tw_data_dec, 2)


        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))
        #Build HTML Table+
        try:
            self.BuildHtml(yamlHandler().read('MemFile'), 0)
            self.BuildHtml(yamlHandler().read('MemFile'), 1)
            self.BuildHtml(yamlHandler().read('MemFile'), 2)
        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))





    def Memory_File(self,data,index):

        #Todo:Anzahl Messmerkmale aus Ini lesen  -> Anz. Spalten prüfen
        self.max_rows= yamlHandler().read("Anz_Zeilen")  #100
        #MEM File Messwerte
        filename=yamlHandler().read('MemFile')[index]
        if not os.path.exists(filename):
            with open(filename,'w') as f:
                if index==0:
                    f.write(";".join(map(str,yamlHandler().read('MW-Header')))+ '\n')
                elif index==1:
                    f.write(";".join(map(str, yamlHandler().read('TW-Header'))) + '\n')
                elif index==2:
                    f.write(";".join(map(str, yamlHandler().read('MW-Header'))) + '\n')
                else:
                    return 0
            #f.close()

        data_str=";".join(map(str,data))

        with open(filename, 'r+') as f:
            content = f.readlines()
            content.insert(1,data_str + '\n')
            while len(content) > (int(self.max_rows + 1)):
                content.pop(-1)

        with open(filename, 'w') as f:
            for line in content:
                f.write(line)


    def MemoryFile_Primary_Setup(self):
        try:
            for index in range(0,len(yamlHandler().read('MemFile'))):
                filename = yamlHandler().read('MemFile')[index]
                if not os.path.exists(filename):
                    with open(filename, 'w') as f:
                        if index == 0:
                            f.write(";".join(map(str, yamlHandler().read('MW-Header'))) + '\n')
                        elif index == 1:
                            f.write(";".join(map(str, yamlHandler().read('TW-Header'))) + '\n')
                        elif index == 2:
                            f.write(";".join(map(str, yamlHandler().read('MW-Header'))) + '\n')
                        else:
                            return 0
                    self.BuildHtml(yamlHandler().read('MemFile'), index)

        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))




    def manipulate_mem_file(self,data):
        #ToDO: Break scheint kompletter Funktion zu escapen -> Fix notwendig
        data = ''.join(data)
        self.filename_mw=yamlHandler().read('MemFile')[0]
        self.filename_tw = yamlHandler().read('MemFile')[1]
        if not os.path.exists(self.filename_mw):
            return 0
        if not os.path.exists(self.filename_tw):
            return 0

        listData = data.split(';')
        self.pwa_id = listData[1][:-1]
        self.MwStatus = int(listData[2], base = 16)
        if self.MwStatus == 0:
            return 0


        try:
            with open(self.filename_mw, 'r+') as f:
                self.content = f.readlines()
                for count,line in enumerate(self.content):
                    if re.search(self.pwa_id,line):
                        self.items = line.split(";")
                        self.items[5] = yamlHandler().read('Status')[self.MwStatus]
                        self.content[count] = ";".join(map(str,self.items))
                        break


            with open(self.filename_mw, 'w') as f:
                for line in self.content:
                    f.write(line)

        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))


        try:
            with open(self.filename_tw, 'r+') as f:
                self.content = f.readlines()
                for count,line in enumerate(self.content):
                    if re.search(self.pwa_id, line):
                        self.items = line.split(";")
                        self.items[5] = yamlHandler().read('Status')[self.MwStatus]
                        self.content[count] = ";".join(map(str, self.items))
                        break

            with open(self.filename_tw, 'w') as f:
                for line in self.content:
                    f.write(line)

        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))




    def Decodieren(self,Eingangs_Block, Laenge_Data):
        #Todo: Catch wrong calls / fill empty fields
        ### PSS Teilessatus ###
        # String zu Hex zu Dez
        try:
            Data = Eingangs_Block

            PSS_Teile_val_str = (Data[60] + Data[61])
            PSS_Teile_value = int(PSS_Teile_val_str, base=16)-1

            #*1:
            try:
                with open(yamlHandler().src_path,'r') as f:
                    self.yaml_cont=yaml.safe_load(f) or {}   # Inhalt der yaml-Dateien in yaml_cont schreiben
            except:
                pass

            ### Seriennummmer ###
            # 2 Werte am Anfang Eingeben, um akt_Zeile als class(List) zu erzeugen
            MW_Zeile = ["'" +
                Data[12] + Data[13] + Data[14] + Data[15] + Data[16] + Data[17] + Data[18] + Data[19] + Data[20] + Data[
                    21] + Data[
                    22] + Data[23] + Data[24], yamlHandler().read('PSS_Teilestatus')[PSS_Teile_value]]      #self.yaml_cont["PSS_TeileStatus"][PSS_Teile_value]]
            TW_Zeile = ["'" +
                Data[12] + Data[13] + Data[14] + Data[15] + Data[16] + Data[17] + Data[18] + Data[19] + Data[20] + Data[
                    21] + Data[
                    22] + Data[23] + Data[24], yamlHandler().read('PSS_Teilestatus')[PSS_Teile_value]]      #self.yaml_cont["PSS_TeileStatus"][PSS_Teile_value]]

            ### Datum und Uhrzeit ###
            MW_Zeile.append(Data[64] + Data[65] + "." + Data[66] + Data[67] + "." + Data[68] + Data[69])  # Datum
            MW_Zeile.append(
                Data[70] + Data[71] + ":" + Data[72] + Data[73] + ":" + Data[74] + Data[
                    75])  # + "," + Data[76] + Data[77] + Data[78])  # Uhrzeit
            TW_Zeile.append(Data[64] + Data[65] + "." + Data[66] + Data[67] + "." + Data[68] + Data[69])  # Datum
            TW_Zeile.append(
                Data[70] + Data[71] + ":" + Data[72] + Data[73] + ":" + Data[74] + Data[
                    75])

            ### Zeitdauer von "Maschine verlassen" bis "Messdaten erhalten" in Millisekunden ###
            # String zu Hex zu Dez
            Traveltime_val_str = (
                        Data[136] + Data[137] + Data[138] + Data[139] + Data[140] + Data[141] + Data[142] + Data[143])
            Traveltime_value = int(Traveltime_val_str, base=16) / 60000  # von ms in min
            Traveltime_value=f"{Traveltime_value:.2f} min"
            MW_Zeile.append(Traveltime_value)
            TW_Zeile.append(Traveltime_value)

            ### Allgemeine Statusangaben ###
            # Messwertstatus
            #Todo: MesswertSta_value -1 ?
            MesswertSta_val_str = (Data[144] + Data[145])
            MesswertSta_value = int(MesswertSta_val_str, base=16)
            #MW_Zeile.append(self.yaml_cont["Status"][MesswertSta_value])   #Messwert_str[MesswertSta_value])
            MW_Zeile.append(yamlHandler().read('Status')[MesswertSta_value])

            # Trendwertstatus
            TrendwertSta_val_str = (Data[146] + Data[147])
            TrendwertSta_value = int(TrendwertSta_val_str, base=16)
            #TW_Zeile.append(self.yaml_cont["Status"][TrendwertSta_value])
            TW_Zeile.append(yamlHandler().read('Status')[TrendwertSta_value])

            # Korrekturen erforderlich
            Korr_feld = (Data[148] + Data[149])
            Korr_feld_bin = bin(int(Korr_feld, base=16)).zfill(8)
            # Messwert Korr. erforderlich
            if Korr_feld_bin[7] == "1":
                MW_Zeile.append("ja")
            else:
                MW_Zeile.append("nein")
            # Trendwert Korr. erforderlich
            if Korr_feld_bin[6] == "1":
                TW_Zeile.append("ja")
            else:
                TW_Zeile.append("nein")
            # Flag Vollkorrektur
            if Korr_feld_bin[5] == "1":
                MW_Zeile.append("ja")
                TW_Zeile.append("ja")
            else:
                MW_Zeile.append("nein")
                TW_Zeile.append("nein")
            # Flag Messdaten verwerfen
            if Korr_feld_bin[4] == "1":
                MW_Zeile.append("ja")
                TW_Zeile.append("ja")
            else:
                MW_Zeile.append("nein")
                TW_Zeile.append("nein")

            ### Merkmale ###
            jj = 0
            for ii in range(152, Laenge_Data + 1, 16):
                jj = jj + 1
                if jj <= 25:  # derzeit 25 Merkmale
                    # Messwert
                    Messwert_val_str = (Data[ii] + Data[ii + 1] + Data[ii + 2] + Data[ii + 3])
                    Messwert_value = int(Messwert_val_str, base=16)
                    if Messwert_value > 32768:  # Vorzeichen beachten
                        Messwert_value = Messwert_value - 65536
                    MW_Zeile.append(Messwert_value)

                    # Messwertstatus
                    MesswertSta_val_str = (Data[ii + 4] + Data[ii + 5])
                    MesswertSta_value = int(MesswertSta_val_str, base=16)
                    #MW_Zeile.append(self.yaml_cont["Status"][MesswertSta_value])#(Messwert_str[MesswertSta_value])
                    MW_Zeile.append(yamlHandler().read('Status')[MesswertSta_value])

                    # Korrekturen erforderlich
                    Korr_feld = (Data[ii + 6] + Data[ii + 7])
                    Korr_feld_bin = (int(Korr_feld, base=2))
                    if Korr_feld_bin == 1:
                        MW_Zeile.append("ja")
                    else:
                        MW_Zeile.append("nein")

                    # Trendwert
                    Trendwert_val_str = (Data[ii + 8] + Data[ii + 9] + Data[ii + 10] + Data[ii + 11])
                    Trendwert_value = int(Trendwert_val_str, base=16)
                    if Trendwert_value > 32768:  # Vorzeichen beachten
                        Trendwert_value = Trendwert_value - 65536
                    TW_Zeile.append(Trendwert_value)

                    # Trendwertstatus
                    TrendwertSta_val_str = (Data[ii + 12] + Data[ii + 13])
                    TrendwertSta_value = int(TrendwertSta_val_str, base=16)
                    #TW_Zeile.append(self.yaml_cont["Status"][TrendwertSta_value])#Messwert_str[TrendwertSta_value]
                    TW_Zeile.append(yamlHandler().read('Status')[TrendwertSta_value])

                    # Korrekturen erforderlich
                    Korr_feld = (Data[ii + 14] + Data[ii + 15])
                    Korr_feld_bin = (int(Korr_feld, base=2))
                    if Korr_feld_bin == 1:
                        TW_Zeile.append("ja")
                    else:
                        TW_Zeile.append("nein")

        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))

        return MW_Zeile,TW_Zeile


    def color_cell(self,val):
        #Lesen aus Mess-/Trendwertstatus + Verworfen durch überholtes Messteil

        # TODO: [gültig: retval[0] = 'background-color: rgba(0,230,64,1)'
        # unültig: retval[0] = 'text-decoration: line-through'
        # Verworfen:background-color: rgba(191,191,191,1)]

        try:
            retval=[]

            srcString = val.loc['MwStatus']

            discarded = re.search('^Verworfen', srcString)
            invalid = re.search('^Ungueltig',srcString)
            io = re.search('^i.O.',srcString)
            for i in range(0,len(val),1):
                retval.append('')
                if discarded:
                    retval[i] = 'background-color: rgba(191,191,191,1) '
                elif invalid:
                    if i > 0:
                        retval[i] = 'background-color: rgba(191,191,191,1)'
                    else:
                        retval[i] = 'background-color: rgba(191,191,191,1)  ; text-decoration: line-through '
                elif io and not i > 0:
                    retval[0] = 'background-color: rgba(0,230,64,1)'
                elif val.iloc[i]=='ja':
                    retval[i-2]='background-color: red'
                elif val.iloc[i]==' X ':
                    retval[i] = 'background-color: red'
            return retval
        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))



    def BuildHtml(self,mem_path,Section):
        try:
            #Section=0 -> MW
            #Section=1 -> TW
            #Section=2 -> ApplCorr
            self.Section=Section
            self.mem_path=(mem_path[self.Section])
            self.dest_path=yamlHandler().read('HTML_Path')[self.Section]
            print(self.mem_path)
            print(self.dest_path)

            ## Header aus csv
            self.data_stream = pd.read_csv(self.mem_path, sep=";", header=0)
            self.data_stream=self.data_stream.replace("'","",regex=True)


            # Set CSS properties for th elements in dataframe
            th_props = [
                ('font-size', '18px'),
                ('text-align', 'center'),
                ('font-weight', 'bold'),
                ('color', '#ffffff'),
                #('color', '#6d6d6d'),
                ('background-color', '#0a408a')
                #('background-color', '#bcb4bs')
            ]

            # Set CSS properties for td elements in dataframe
            td_props = [
                ('font-size', '16px'),
                ('text-align', 'center'),
                ('white-space', 'nowrap'),
                ('overflow', 'hidden'),
                ('border-bottom','1px solid #0a408a')
            ]


            # Set table styles
            styles = [
                dict(selector="th", props=th_props),
                dict(selector="td", props=td_props)
                #dict(selector="tr:hover", props=[('background-color','orange')]) #,  td:hover
                #dict(selector="tr:nth-child(even)", props=[('background-color', 'f2f2f2')])
            ]

            if self.Section==0 or self.Section==2:
                self.data_stream.reindex(yamlHandler().read("MW-Header"), axis="columns")
            if self.Section==1:
                self.data_stream.reindex(yamlHandler().read("TW-Header"), axis="columns")



            overhead=self.data_stream.iloc[:, 0:6]
            #Out of correction limit general
            oocl_general=self.data_stream.iloc[:,6].copy()
            #toCopy Workaround for "Setting With Copy Warning"
            oocl_general[oocl_general != "ja"] = " O "
            oocl_general=oocl_general.replace("ja"," X ")
            oocl_general=pd.Series(oocl_general).rename("out of corr. lim")
            values=self.data_stream[self.data_stream.columns[9:]].copy()
            self.table=pd.concat([overhead,oocl_general,values],axis=1)


            self.hidden_cols=[]
            self.hidden_cols.extend(["Traveltime"])
            for i in range(1,26):
                if self.Section==0:
                    self.hidden_cols.extend(["Mw"+str(i)+"Status","Korr."+str(i)])
                if self.Section==1:
                    self.hidden_cols.extend(["Tw"+str(i)+"Status","Korr."+str(i)])
                if self.Section==2:
                    self.hidden_cols.extend(["Mw"+str(i)+"Status","Korr."+str(i)])

            self.data_stream = (self.table
                  .style
                  .apply(self.color_cell,axis=1)
                  .hide_columns(self.hidden_cols)
                  .hide_index())
            self.data_stream=self.data_stream.set_table_styles(styles)

            self.html = self.data_stream.render()
            with open(self.dest_path,'w') as file:
                file.write(self.html)

            print("New table rendered from " + self.mem_path + " :")
            print(overhead.iloc[:1,:3])
            return

        except:
            print(sys.exc_info()[1])
            print("Error on line {}".format(sys.exc_info()[-1].tb_lineno))



if __name__ == "__main__":

##Param: Path,Anz. Spalten/Zeilen, Filename
    yamlHandler().prim_setup()
    #For local Test
    if not getattr(sys,'frozen',False):
        yamlHandler().add('plcPath','C:\\temp\\')
        yamlHandler().add('HTML_Path',['C:\\temp\\MW_Site.html','C:\\temp\\TW_Site.html','C:\\temp\\Appl_Site.html'])
    plcEventHandler().MemoryFile_Primary_Setup()
    plcWatcher(yamlHandler().read('plcPath')).run()

    #plcEventHandler().BuildHtml(path,dest_path)
    #path="C:\\temp"
    #plcWatcher(path).run()
