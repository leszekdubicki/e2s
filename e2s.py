# -*- coding: cp1250 -*-

#programik do przerzucania wyników e-maxa do otwartego szkicu solida:

from win32com.client import gencache, constants, pythoncom
import win32com.client          #SolidWorks type library
import sys, os
import os.path
import getopt
import e_max
import win32api, time

selectionlist = e_max.selectionlist; get_planes = e_max.get_planes

#usage notify:
USER = win32api.GetUserName()
#if os.path.isfile("F:\Common\MDC\PL\General\Medical Engineering\DEPT\Tools for SolidWorks\macros\log"):
#	try:
#		FILE = open("F:\Common\MDC\PL\General\Medical Engineering\DEPT\Tools for SolidWorks\macros\log","a")
#		FILE.write(sys.argv[0] + " " + USER + time.strftime(" %Y-%m-%d %H:%M:%S") + "\n")
#		FILE.close()
#	except:
#		A = None	#blank line 
SOFT = 1
if __name__ == "__main__":
	ARGS = sys.argv

	OPTS, ARGS = getopt.getopt(ARGS[1:],'ps')
	#print print
	MEAS_POINTS = 0; MULTISKETCH = 0
	for O in OPTS:
		if '-p' in O:
			MEAS_POINTS = 1
		if '-s' in O:
			MULTISKETCH = 1

	#argumenty to pliki:
	FILES = []
	if len(ARGS)>0:
		for A in ARGS:
			if len(A)<4:
				FILES.append(A+'.tch')
			else:
				#weryfikacja, czy mamy plik czy katalog:
				if os.path.isdir(A):
					#wyciągnięcie wszystkich nazw plików z katalogu:
					FILES_FROM_DIR = os.listdir(A)
					#print FILES_FROM_DIR
					for F in FILES_FROM_DIR:
						if os.path.isfile(A+'\\'+F):
							FILES.append(A+'\\'+F)
				else:
					if A[-4:].lower()=='.tch':
						FILES.append(A)
	else:
		#poproszenie o podanie pliku:
		import Tkinter as tk
		import tkFileDialog
		root = tk.Tk()
		root.withdraw()
		FILE = tkFileDialog.askopenfilename()
		if len(FILE)==0:
			print("brak pliku pmiarowego")
			raw_input()
			sys.exit()
		FILES.append(FILE)
		#print "brak pliku"
		#raw_input()
	
	
	#rozdzial na Solida i PROE:
	#Zmienna SOFT:
	#	0 - Solid
	#	1 - eksport do DXF
	
	if SOFT == 0:
		sldmod = gencache.EnsureModule('{83A33D31-27C5-11CE-BFD4-00400513BB57}', 0, 17, 0)

		swx = win32com.client.Dispatch("SldWorks.Application")        # getinstance of SolidWorks
		swx = sldmod.ISldWorks(swx)

		adoc = swx.ActiveDoc            # get active document
		if adoc==None:
			TEMLPATE = "F:\\Common\\MDC\\PL\\General\\Medical Engineering\\DEPT\\Tools for SolidWorks\\PL-MDC-Template-sld-drawings\\Part.PRTDOT"
			if os.path.isfile(TEMLPATE):
				print(u"Otwieram Standardowy tamplate części...")
				#swx.OpenDoc6(TEMLPATE,0x1,0x0,None,None,None)
				adoc = swx.NewDocument(TEMLPATE,0,0,0)
				#adoc = swx.ActiveDoc            # get active document
			else:
				sys.exit("brak pliku SolidWorks")
		adoc = sldmod.IModelDoc2(adoc)            # get active document

		sk = adoc.GetActiveSketch2()              # get active sketch
		if sk==None:
			s = adoc.SelectionManager
			SELLIST = selectionlist(s)
			PLANES = get_planes(SELLIST)
			if len(PLANES)==0:
				#to bierzemy płaszczyznę przednią
				adoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, None, 0)
				SELLIST = selectionlist(s)
				PLANES = get_planes(SELLIST)
			PLANE = PLANES[0]
			adoc.InsertSketch2(1)
			sk = adoc.GetActiveSketch2()
		#sk = sldmod.ISketch
	
		oldAutoSolve = sk.GetAutomaticSolve
		sk.SetAutomaticSolve(0)         # turn off auto solve
		adoc.SetAddToDB(1)              # add ents directly todatabase
		adoc.SetDisplayWhenAdded(0)     # don't display ents untilredraw
	
		for F in FILES:
			if (MULTISKETCH) and (not F==FILES[0]):
				#utworzenie nowego szkicu
				adoc.Extension.SelectByID2(PLANE.Name, "PLANE", 0, 0, 0, False, 0, None, 0)
				adoc.InsertSketch2(1)
				sk = adoc.GetActiveSketch2()
				oldAutoSolve = sk.GetAutomaticSolve
		
			if (F[len(F)-4:]).lower()=='.tch':
				TCHFILE = e_max.tch_file(F)
			else:
				TCHFILE = e_max.prg_file(F)
			TCHFILE.get_objects()
			for i in range(0,len(TCHFILE._objects)):
				OBJ = TCHFILE._objects[i]
				if OBJ==None:
					continue
				if OBJ.type()=='POINT':
					p=OBJ.P()
					ent = adoc.CreatePoint2(p[0]/1000., p[1]/1000., p[2]/1000.)
				elif OBJ.type()=='LINE':
					p1=OBJ.P1(); p2=OBJ.P2()
					ent = adoc.CreateLine2(p1[0]/1000., p1[1]/1000., p1[2]/1000., p2[0]/1000., p2[1]/1000., p2[2]/1000.)
				elif OBJ.type()=='CIRCLE':
					p1=OBJ.O(); p2=[p1[0]+OBJ.d()/2.,p1[1],p1[2]]
					ent = adoc.CreateCircle2(p1[0]/1000., p1[1]/1000., p1[2]/1000., p2[0]/1000., p2[1]/1000., p2[2]/1000.)
	
				if MEAS_POINTS:
					PTS = OBJ.meas_points()
					for P in PTS:
						ent = adoc.CreatePoint2(P[0]/1000., P[1]/1000., P[2]/1000.)
			if MULTISKETCH:
				#zmiana nazwy szkicu na nazwe pliku i zamknięcie szkicu
				sk.Name = F
				adoc.InsertSketch2(1)
				sk.SetAutomaticSolve(oldAutoSolve)      # restore autosolve
						
		
		

		sk.SetAutomaticSolve(oldAutoSolve)      # restore autosolve
		adoc.SetDisplayWhenAdded(1)
		adoc.SetAddToDB(0)
		adoc.GraphicsRedraw2()            #show us the new slots!
	elif SOFT == 1:
		#to DXF
		for F in FILES:
			#katalog do zapisu na dxf:
			PATH = os.path.dirname(F)
			FILENAME = os.path.basename(F)[:-3]+"dxf"
			DXF_FILE_NAME = PATH+"\\"+FILENAME
			E = e_max.tch_file(F)
			E.get_objects()
			E.converttodxf(DXF_FILE_NAME)
			#miało być creo vbapi ale zrezygnowałem:
			#creomod = gencache.EnsureModule('{176453F2-6934-4304-8C9D-126D98C1700E}', 0, 1, 0)
			#crx = win32com.client.Dispatch(".Application")        # getinstance of SolidWorks 
	
