#!/usr/bin/env python
__author__ = 'Matthew Loschmann'
#email: matt.loschmann@gmail.com

import gi
import json
import pymssql
import os
import subprocess #for calling qrcode generator and glabel
import shutil
import unicodecsv as csv
import xml.etree.ElementTree as ET
import ezodf
import zipfile
import tempfile
import platform
import sys

#Load GTK
gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, Gdk, GdkPixbuf

#Global variables


#Classes
class Case:
    def __init__(self, location, occurrence_number, evidence_number, pieces_of_evidence, submitted_date,
                     investigating_officer):
        self.location = location
        self.occurrence_number = occurrence_number
        self.evidence_number = evidence_number
        self.pieces_of_evidence = pieces_of_evidence
        self.submitted_date = submitted_date
        self.investigating_officer = investigating_officer

class Handler:
    def onDeleteWindow(self, *args):
        Gtk.main_quit(*args)

    def onButtonPressed(self, cancel_btn):
        Gtk.main_quit()

    def on_save_settings_pressed(self, save_settings):
        s_settings["XWAYS"] = xw_chk.get_active()
        s_settings["ENCASE"] = enc_chk.get_active()
        s_settings["IEF"] = ief_chk.get_active()
        s_settings["AXIOM"] = ax_chk.get_active()
        s_settings["CELLEBRITE"] = cel_chk.get_active()
        s_settings["AUTOPSY"] = aut_chk.get_active()
        s_settings["WARRANT_PHOTOS"] = war_chk.get_active()
        s_settings["COPY_DISCLOSURE_MEMO"] = cpdis_chk.get_active()
        s_settings["DISCLOSURE_MEMO"] = dmfc.get_filename()
        s_settings["FINAL_DISCLOSURE"] = fin_dis.get_active()
        s_settings["FORENSIC_IMAGES"] = for_imgs.get_active()
        s_settings["REPORTS"] = reports.get_active()
        s_settings["EXPORT"] = export_folder.get_active()
        s_settings["DEVICE_PHOTOS"] = dev_photos.get_active()
        s_settings["SERVER1"] = srv_one.get_filename()
        s_settings["SERVER2"] = srv_two.get_filename()
        s_settings["OTHER_LOC"] = srv_oth.get_filename()
        s_settings["USER_INITIALS"] = inv_init.get_text()

        with open("setup_settings.json","w") as settings_file:
            json.dump(s_settings,settings_file)
        settings_file.close()

    def on_label_only_pressed(self,label_only):
        print "print labels only"

        srv_location = ""
        inc_num = occNumEntry.get_text()
        inc_num = inc_num + "_" + inv_init.get_text()
        c = Case(srv_location, inc_num, EvNumEntry.get_text(), EvItemCnt.get_text(), SubDate.get_text(),
                 InvOff.get_text())
        ##generate labels
        generate_labels(c, True)

    def on_setup_btn_pressed(self,setup_btn):
        print "Full Case Setup"
        case_setup()

    def on_occ_search_db_btn_pressed(self,occ_search_db_btn):
        Incident_Num = occNumEntry.get_text()
        SQL_SEL = "SELECT * FROM Occurrence O " \
                  "LEFT JOIN CFOccurrence CF ON O.CyberNumber = CF.CyberNumber " \
                  "LEFT JOIN Employee E ON E.EmployeeID = O.InvestigatorID " \
                  "WHERE O.OccurrenceNumber like \'%" + Incident_Num + "%\'"

        cursor.execute(SQL_SEL)
        row = cursor.fetchone()
        rec_date = row["ReceivedDate"]
        SubDate.set_text(rec_date[:16])
        EvNumEntry.set_text(row['CyberNumber'])
        EvItemCnt.set_text(str(row['Exhibits']))
        InvOff.set_text(row['LastName'])

    def on_tcu_search_db_btn_pressed(self,tcu_search_db_btn):
        CyberUnitID = EvNumEntry.get_text()
        SQL_SEL = "SELECT * FROM Occurrence O " \
                  "LEFT JOIN CFOccurrence CF ON O.CyberNumber = CF.CyberNumber " \
                  "LEFT JOIN Employee E ON E.EmployeeID = O.InvestigatorID " \
                  "WHERE O.CyberNumber like \'%" + CyberUnitID + "%\'"

        cursor.execute(SQL_SEL)
        row = cursor.fetchone()
        rec_date = row["ReceivedDate"]
        SubDate.set_text(rec_date[:16])
        # SubmitDateEntry.delete(16, 'end')
        occNumEntry.set_text(row['OccurrenceNumber'])
        EvItemCnt.set_text(str(row['Exhibits']))
        InvOff.set_text(row['LastName'])

    def on_cpy_disclosure_memo_checkbutton_toggled(self,cpy_disclosure_memo_checkbox):
        if cpdis_chk.get_active():
            dmfc.show()
        else:
            dmfc.hide()

#Functions
def load_settings():
    with open("setup_settings.json") as settings_file:
    # load saved settings
        saved_settings = json.load(settings_file)
    # Close settings file
    settings_file.close()

    if saved_settings["XWAYS"] == True:
        xw_chk.set_active(True)
    else:
        xw_chk.set_active(False)

    if saved_settings["ENCASE"] == True:
        enc_chk.set_active(True)
    else:
        enc_chk.set_active(False)

    if saved_settings["IEF"] == True:
        ief_chk.set_active(True)
    else:
        ief_chk.set_active(False)

    if saved_settings["AXIOM"] == True:
        ax_chk.set_active(True)
    else:
        ax_chk.set_active(False)

    if saved_settings["CELLEBRITE"] == True:
        cel_chk.set_active(True)
    else:
        cel_chk.set_active(False)
        
    if saved_settings["AUTOPSY"] == True:
        aut_chk.set_active(True)
    else:
        aut_chk.set_active(False)

    if saved_settings["WARRANT_PHOTOS"] == True:
        war_chk.set_active(True)
    else:
        war_chk.set_active(False)

    if saved_settings["COPY_DISCLOSURE_MEMO"] == True:
        cpdis_chk.set_active(True)
        dmfc.show()
    else:
        cpdis_chk.set_active(False)
        dmfc.hide()

    dmfc.set_filename(saved_settings["DISCLOSURE_MEMO"])

    if saved_settings["FINAL_DISCLOSURE"] == True:
        fin_dis.set_active(True)
    else:
        fin_dis.set_active(False)

    if saved_settings["FORENSIC_IMAGES"] == True:
        for_imgs.set_active(True)
    else:
        for_imgs.set_active(False)

    if saved_settings["REPORTS"] == True:
        reports.set_active(True)
    else:
        reports.set_active(False)

    if saved_settings["EXPORT"] == True:
        export_folder.set_active(True)
    else:
        export_folder.set_active(False)

    if saved_settings["DEVICE_PHOTOS"] == True:
        dev_photos.set_active(True)
    else:
        dev_photos.set_active(False)

    srv_one.set_filename(saved_settings["SERVER1"])
    srv_two.set_filename(saved_settings["SERVER2"])
    srv_oth.set_filename(saved_settings["OTHER_LOC"])
    inv_init.set_text(saved_settings["USER_INITIALS"])
    return saved_settings

def srv_connect():
    # Open Server settings
    with open("server_cnct.json") as server_file:
    # load server settings
        sql_server_connect = json.load(server_file)
    # Close server settings file
    server_file.close()
    return sql_server_connect

def update_disclosure(cc):
    count = 0
    if fin_dis.get_active() == True:
        path = cc.location + "/" + cc.occurrence_number + "/Final_Disclosure/"
    else:
        path = cc.location + "/" + cc.occurrence_number + "/"

    doc = ezodf.opendoc(path + 'Disclosure_Memo.odt')
    if doc.doctype == 'odt':
        pass
        print "Disclosure_Memo.odt found."

    a = zipfile.ZipFile(path + 'Disclosure_Memo.odt')
    zipfile.ZipFile.extractall(a, "temp")

    ###Update Content.xml
    print "Creating Temporary content.xml"
    with open("temp/content-temp.xml", "wt") as fout:
        with open("temp/content.xml", "rt") as fin:
            print "Replacing Incident Number in Body"
            for line in fin:
                fout.write(line.replace('XXINC#XX', cc.occurrence_number[:11]))
            count = count + 1
    print "Replacing Content.xml"
    os.remove("temp/content.xml")
    os.rename("temp/content-temp.xml", "temp/content.xml")
    print str(count) + " instances of XXINC#XX replaced with " + cc.occurrence_number[:11]

    # reset counter
    count = 0

    ###    Update Styles.xml

    print "Creating Temporary styles.xml"
    with open("temp/styles-temp.xml", "wt") as fout:
        with open("temp/styles.xml", "rt") as fin:
            print "Replacing Incident Number in Header/Footer"
            for line in fin:
                fout.write(line.replace('XXINC#XX', cc.occurrence_number[:11]))
            count = count + 1
    print "Replacing styles.xml"
    os.remove("temp/styles.xml")
    os.rename("temp/styles-temp.xml", "temp/styles.xml")
    print str(count) + " instances of XXINC#XX replaced with " + cc.occurrence_number[:11]

    ###Create New Disclosure Memo
    zf = zipfile.ZipFile(path + "Disclosure_Memo2.odt", "w")
    for dirname, subdirs, files in os.walk("temp"):
        # zf.write(dirname)
        for filename in files:
            newpath = os.path.join(dirname, filename)
            zf.write(newpath, newpath.replace("temp/", ""))
    zf.close()

    print "Cleaning up..."
    shutil.rmtree("temp/")
    print "removing" + path + 'Disclosure_Memo.odt'
    os.remove(path + 'Disclosure_Memo.odt')
    print "renaming" + path + 'Disclosure_Memo2.odt'
    os.rename(path + 'Disclosure_Memo2.odt', path + 'Disclosure_Memo.odt')

def generate_labels(case_label, only=False):
    if only == False:
        if sys.platform == "linux" or sys.platform == "linux2":
            path = "{}/{}/".format(case_label.location, case_label.occurrence_number)
        elif sys.platform == "wind32":
            path = "{}\\{}\\".format(case_label.location, case_label.occurrence_number)

        with open(path + 'labels.csv', 'wb') as csvfile:
            labelwriter = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            labelwriter.writerow(['inc_num'] + ['exhibit'] + ['investigator_name'] + ['submit_date'])
            labelwriter.writerow(
                [case_label.occurrence_number] + [case_label.evidence_number] + [case_label.investigating_officer] + [
                    case_label.submitted_date])
            labelwriter.writerow(
                [case_label.occurrence_number] + [case_label.evidence_number] + [case_label.investigating_officer] + [
                    case_label.submitted_date])
            labelwriter.writerow(
                [case_label.occurrence_number] + [case_label.evidence_number] + [case_label.investigating_officer] + [
                    case_label.submitted_date])
            for i in xrange(1, int(case_label.pieces_of_evidence) + 1):
                labelwriter.writerow([case_label.occurrence_number] + [case_label.evidence_number + "-{}".format(i)] + [
                    case_label.investigating_officer] + [case_label.submitted_date])
        csvfile.close()
        ET.register_namespace('', "http://glabels.org/xmlns/3.0/")
        tree = ET.parse("Evidence_Label_Template.xml")
        root = tree.getroot()
        root[2].set('src',path + "labels.csv")
        tree.write(path + case_label.occurrence_number + ".glabels")
        subprocess.call(["glabels-3", path + case_label.occurrence_number + ".glabels"])
    else:
        with open("./labels/" + case_label.occurrence_number + ".csv", 'wb') as csvfile:
            labelwriter = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            labelwriter.writerow(['inc_num'] + ['exhibit'] + ['investigator_name'] + ['submit_date'])
            labelwriter.writerow(
                [case_label.occurrence_number] + [case_label.evidence_number] + [case_label.investigating_officer] + [
                    case_label.submitted_date])
            labelwriter.writerow(
                [case_label.occurrence_number] + [case_label.evidence_number] + [case_label.investigating_officer] + [
                    case_label.submitted_date])
            labelwriter.writerow(
                [case_label.occurrence_number] + [case_label.evidence_number] + [case_label.investigating_officer] + [
                    case_label.submitted_date])
            for i in xrange(1, int(case_label.pieces_of_evidence) + 1):
                labelwriter.writerow([case_label.occurrence_number] + [case_label.evidence_number + "-{}".format(i)] + [
                    case_label.investigating_officer] + [case_label.submitted_date])
        csvfile.close()
        ET.register_namespace('', "http://glabels.org/xmlns/3.0/")
        tree = ET.parse("Evidence_Label_Template.xml")
        root = tree.getroot()
        root[2].set('src', "./labels/" + case_label.occurrence_number + ".csv")
        tree.write("./labels/" + case_label.occurrence_number + ".glabels")
        subprocess.call(["glabels-3", "./labels/" + case_label.occurrence_number + ".glabels"])

def case_setup():
    srv_location = ""
    inc_num  = occNumEntry.get_text()
    inc_num = inc_num + "_" + inv_init.get_text()

    if sys.platform == "linux" or sys.platform == "linux2":
        slash = "/"
    elif sys.platform == "wind32":
        slash = "\\"

    #Get base path of case location
    if srv_one_radio.get_active() == True:
        srv_location = srv_one.get_filename()
    elif srv_two_radio.get_active() == True:
        srv_location = srv_two.get_filename()
    elif srv_oth_radio.get_active() == True:
        srv_location = srv_oth.get_filename()

    if sys.platform == "linux" or sys.platform == "linux2":
        print "Linux OS Detected"
        path = "{}/{}/".format(srv_location,inc_num)
    elif sys.platform == "wind32":
        print "Windows OS Detected"
        path = "{}\\{}\\".format(srv_location, inc_num)

    c  = Case(srv_location, inc_num, EvNumEntry.get_text(), EvItemCnt.get_text(), SubDate.get_text(), InvOff.get_text())

    #Create base folder
    try:
        os.makedirs(path)
    except:
        print "Base Case Folder already exists"

    if xw_chk.get_active() == True:
        try:
            os.makedirs(path + "XWAYS")
        except:
            print "Unable to create folder XWAYS"

    if enc_chk.get_active() == True:
        try:
            os.makedirs(path + "EnCase")
        except:
            print "Unable to create folder EnCase"

    if ief_chk.get_active() == True:
        try:
            os.makedirs(path + "IEF")
        except:
            print "Unable to create folder IEF"
            
    if aut_chk.get_active() == True:
        try:
            os.makedirs(path + "Autopsy")
        except:
            print "Unable to create folder Autopsy"

    if ax_chk.get_active() == True:
        try:
            os.makedirs(path + "AXIOM")
        except:
            print "Unable to create folder AXIOM"

    if cel_chk.get_active() == True:
        try:
            os.makedirs(path + "Cellebrite")
        except:
            print "Unable to create folder Cellebrite"

    if war_chk.get_active() == True:
        try:
            os.makedirs(path + "Warrant_Photos")
        except:
            print "Unable to create folder Warrant_Photos"

    if fin_dis.get_active() == True:
        try:
            os.makedirs(path + "Final_Disclosure")
            os.makedirs(path + "Final_Disclosure{}Support_Files".format(slash))
        except:
            print "Unable to create folder Final_Disclosure"

    if cpdis_chk.get_active() == True & fin_dis.get_active() == True:
        try:
            shutil.copy(dmfc.get_filename(), path + "Final_Disclosure")
        except:
            print "Unable to copy {} to {} /Final_Disclosure/".format(dmfc.get_filename, srv_location)
    elif cpdis_chk.get_active() == True:
        try:
            os.makedirs(path + "Support_Files")
            shutil.copy(dmfc.get_filename(), srv_location + "/" + inc_num)
        except:
            print "Unable to copy {} to {}/".format(dmfc.get_filename, srv_location)

    if EvNumEntry.get_text():
        for i in range (1, int(EvItemCnt.get_text())+1):
            try:
                os.makedirs(path + EvNumEntry.get_text() + "-{}".format(i))

                if for_imgs.get_active() == True:
                    os.makedirs(path + EvNumEntry.get_text() + "-{}{}Forensic_Images".format(i,slash))

                if reports.get_active() == True:
                    os.makedirs(path + EvNumEntry.get_text() + "-{}{}Reports".format(i,slash))

                if export_folder.get_active() == True:
                    os.makedirs(path + EvNumEntry.get_text() + "-{}{}Export".format(i,slash))

                if dev_photos.get_active() == True:
                    os.makedirs(path + EvNumEntry.get_text() + "-{}{}Device_Photos".format(i,slash))
            except:
                print "Could not create Exhibit folder"

    if cpdis_chk.get_active() == True:
        update_disclosure(c)

    if sys.platform == "linux" or sys.platform == "linux2":
        qrcode_gen = inc_num + " " + EvNumEntry.get_text() + " " \
                     + EvItemCnt.get_text() + " " \
                     + SubDate.get_text() + " " + InvOff.get_text()

        subprocess.check_call(['xdg-open', path])
        ##Create QRCODE
        subprocess.call(["qrencode", "-o", path + "qrcode.png", qrcode_gen])
        qr_code.set_from_pixbuf(GdkPixbuf.Pixbuf.new_from_file(path + "qrcode.png"))
        ##generate labels
        generate_labels(c)


#Load GUI Template
builder = Gtk.Builder()
builder.add_from_file("Case_setup.glade")
builder.connect_signals(Handler())

srv_one_radio = builder.get_object("server1_btn")
srv_two_radio = builder.get_object("server2_btn")
srv_oth_radio = builder.get_object("other_btn")

qr_code = builder.get_object("qr_image")

xw_chk = builder.get_object("xways_checkbutton")
enc_chk = builder.get_object("encase_checkbutton")
ief_chk = builder.get_object("ief_checkbutton")
ax_chk = builder.get_object("axiom_checkbutton")
cel_chk = builder.get_object("cellebrite_checkbutton")
aut_chk = builder.get_object("autopsy_checkbutton")
war_chk = builder.get_object("warrant_checkbutton")
cpdis_chk = builder.get_object("cpy_disclosure_memo_checkbutton")
dmfc = builder.get_object("disclosure_memo_filechooser")
fin_dis = builder.get_object("final_disclosure_checkbutton")
for_imgs = builder.get_object("forensic_imgs_checkbutton")
reports = builder.get_object("reports_checkbutton")
export_folder = builder.get_object("export_checkbutton")
dev_photos = builder.get_object("device_photos_checkbutton")
srv_one = builder.get_object("server1filechooser")
srv_two = builder.get_object("server2filechooser")
srv_oth = builder.get_object("otherfilechooser")
inv_init = builder.get_object("Inv_Initials")

EvNumEntry = builder.get_object("tcu_entry")
occNumEntry = builder.get_object("occ_entry")
EvItemCnt = builder.get_object("num_items_entry")
SubDate = builder.get_object("submit_date_entry")
InvOff = builder.get_object("inv_officer_entry")
occ_search = builder.get_object("occ_search_db_btn")
tcu_search = builder.get_object("tcu_search_db_btn")
lbl_only = builder.get_object("label_only")

if sys.platform == "linux" or sys.platform == "linux2":
    qr_code.hide()
    lbl_only.hide()

#Show window
window = builder.get_object("case_setup")
window.show_all()

if __name__ == "__main__":

    #load saved settings
    s_settings = load_settings()

    #load server connection settings
    sql_server_connect = srv_connect()

    #Connect to the server
    conn = pymssql.connect(sql_server_connect["host"], sql_server_connect["username"], sql_server_connect["password"], sql_server_connect["database"])
    cursor = conn.cursor(as_dict=True)

    #Main application Window
    Gtk.main()

    #Close sqlite connection
    cursor.close()
    conn.close()
