# -*- coding: cp1250 -*-

#program do odczytywania danych z plików utworzonych przez E-max (pliki *.tch) i wstawiania zmierzonych punktów do solidworksa

from copy import deepcopy; cp = deepcopy
def selectionlist (selmgr):
        """Returns list of selected objects and their types{'obj':xxx, 'type':nn}"""
        l = []
        for i in range(1,selmgr.GetSelectedObjectCount+1):
                o = selmgr.GetSelectedObject(i)
                t = selmgr.GetSelectedObjectType(i)
                l.append({"obj":o, "type":t})
        return l

def get_planes (sellist):
        clineset = []
        for i in sellist:
                etype = i['type']
                if (etype == 4):       #sketch segment
                        ent = i['obj']
                        clineset.append(ent)
        return clineset

class tch_file:
	def __init__(self,nazwa_pliku):
		PLIK=open(nazwa_pliku,'r')
		self._text=PLIK.readlines()
		self._nazwa_pliku = nazwa_pliku
		if '\\' in self._nazwa_pliku:
			self._path = self._nazwa_pliku[0:self._nazwa_pliku.rindex('\\')]
			self._file_name = self._nazwa_pliku[self._nazwa_pliku.rindex('\\')+1:]
		else:
			self._path = '.'
		PLIK.close()
		self._objects=[]
	def get_objects(self):
		#analizuje plik pod kątem zmierzonych obiektów:
		i=0
		while i in range(0,len(self._text)):
			#L=self.split_line(i)
			OBJ,i=self.recognize_obj(i)
			self._objects.append(OBJ)
			i+=1
		
	def split_line(self,line):
		#rozbija linie na elementy wg schematu:
		#	<OBJECT (NR)> = COMMANT / <PARAMETRY>
		L=self._text[line]
		LINIA={'pos':line}
		if '=' in L:
			ind_nawias_1=L.index('(')
			ind_nawias_2=L.index(')')
			LINIA['obj_type']=L[0:ind_nawias_1]
			LINIA['obj_no']=int(L[ind_nawias_1+1:ind_nawias_2])
			ind_eq=L.index('=')
		else:
			LINIA['obj_type']=None
			LINIA['obj_no']=None
			ind_eq=-1
		if '/' in L:
			ind_sl=L.index('/')
			LINIA['command']=(L[ind_eq+1:ind_sl]).strip()
			LINIA['params']=[]
			if len(L)>ind_sl+1:
				PARAMETRY=L[ind_sl+1:]
				PARAMETRY = PARAMETRY.replace(',',' ')
				PARAMETRY=PARAMETRY.split()
				for P in PARAMETRY:
					LINIA['params'].append(P)
		else:
			LINIA['command']=L[ind_eq+1:].strip()
			LINIA['params']=[]
			
		return LINIA
	def recognize_obj(self,NR_BLOKU):
		#rozpoznaje komende i interpretuje text:
		i=NR_BLOKU; BLOK=self.split_line(i)
		OBJECT = None
		#print BLOK
		if BLOK['command']=='MEAS':
			#trzeba znaleźć zakres bloków odpowiadających temu pomiarowi:
			i=BLOK['pos']+1
			B=self.split_line(i)
			MEAS_BLOCKS = [BLOK,B]
			while not B['command']=='ENDMEAS':
				i+=1
				B=self.split_line(i)
				MEAS_BLOCKS.append(B)
			i += 1
			MEAS_BLOCKS.append(self.split_line(i))	#dołączenie bloku z parametrami obiektu
			if BLOK['params'][0]=='POINT':
				OBJECT = point(MEAS_BLOCKS)
			elif BLOK['params'][0]=='CIRCLE':
				OBJECT = circle(MEAS_BLOCKS)
			elif BLOK['params'][0]=='ELLIPSE':
				OBJECT = ellipse(MEAS_BLOCKS)
			elif BLOK['params'][0]=='LINE':
				OBJECT = line(MEAS_BLOCKS)
			#print OBJECT.meas_points()
		return OBJECT, i
	def fileformat(self):
		#sprawdzenie rozszerzenia pliku:
		EXT = self._nazwa_pliku[len(self._nazwa_pliku)-3:]
		if EXT.lower() == 'tch':
			FORMAT = 1	#e-max
		elif EXT.lower() == 'txt':
			FORMAT = 2	#flex---
		else:
			FORMAT = 0	#nieznany format
	def converttodxf(self, nazwapliku = None, points = True, d3d = True):
		#konwertuje uzywajac modulu dxfwrite
		#points - czy wstawiac punkty wybierane podczas montazu
		if nazwapliku == None:
			nazwapliku = os.path.basename(self._nazwa_pliku)[:-3]+"dxf"
		from dxfwrite import DXFEngine as dxf
		drawing = dxf.drawing(nazwapliku)
		for OBJ in self._objects:
			#print "hehe"
			if OBJ==None:
				continue
			if OBJ.type()=='POINT':
				p=OBJ.P()
				if d3d:
					drawing.add(dxf.point((p[0],p[1],p[2])))
				else:
					drawing.add(dxf.point((p[0],p[1])))
			elif OBJ.type()=='LINE':
				p1=OBJ.P1(); p2=OBJ.P2()
				drawing.add(dxf.line((p1[0], p1[1]), (p2[0], p2[1])))
				#ent = adoc.CreateLine2(p1[0]/1000., p1[1]/1000., p1[2]/1000., p2[0]/1000., p2[1]/1000., p2[2]/1000.)
			elif OBJ.type()=='CIRCLE':
				p1=OBJ.O()
				drawing.add(dxf.circle(OBJ.d()/2.,(p1[0], p1[1])))
				#ent = adoc.CreateCircle2(p1[0]/1000., p1[1]/1000., p1[2]/1000., p2[0]/1000., p2[1]/1000., p2[2]/1000.)
			if points and not OBJ.type()=="POINT":
				PTS = OBJ.meas_points()
				print "dodawanie punktow pomiarowych:"
				for pp in PTS:
					#print (pp[0],pp[1],pp[2])
					if d3d:
						drawing.add(dxf.point((pp[0],pp[1],pp[2])))
					else:
						drawing.add(dxf.point((pp[0],pp[1])))
					
		drawing.save()

	def solidsketch(self, points = True, d3d = False):
		from win32com.client import gencache, constants, pythoncom
		import win32com.client          #SolidWorks type library
		sldmod = gencache.EnsureModule('{83A33D31-27C5-11CE-BFD4-00400513BB57}', 0, 19, 0)

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
		for OBJ in self._objects:
			if OBJ==None:
				continue
			if OBJ.type()=='POINT':
				p=OBJ.P()
				if d3d:
					z = p[2]/1000.
				else:
					z = 0.
				ent = adoc.CreatePoint2(p[0]/1000., p[1]/1000., z)
			elif OBJ.type()=='LINE':
				p1=OBJ.P1(); p2=OBJ.P2()
				ent = adoc.CreateLine2(p1[0]/1000., p1[1]/1000., p1[2]/1000., p2[0]/1000., p2[1]/1000., p2[2]/1000.)
			elif OBJ.type()=='CIRCLE':
				p1=OBJ.O(); p2=[p1[0]+OBJ.d()/2.,p1[1],p1[2]]
				ent = adoc.CreateCircle2(p1[0]/1000., p1[1]/1000., p1[2]/1000., p2[0]/1000., p2[1]/1000., p2[2]/1000.)
			if points:
				PTS = OBJ.meas_points()
				for P in PTS:
					if d3d:
						z = P[2]/1000.
					else:
						z = 0.
					ent = adoc.CreatePoint2(P[0]/1000., P[1]/1000., z)
		sk.SetAutomaticSolve(oldAutoSolve)      # restore autosolve
		adoc.SetDisplayWhenAdded(1)
		adoc.SetAddToDB(0)
		adoc.GraphicsRedraw2()            #show us the new slots!
			
		
		

class prg_file(tch_file):
	def __init__(self,nazwa_pliku):
		tch_file.__init__(self,nazwa_pliku)
		for i in range(0,len(self._text)):
			if 'HEADEREND' in self._text[i]:
				HEADEND = cp(i)
			elif 'STATUSEND' in self._text[i]:
				STATUSEND = cp(i)
			elif 'PRG-STEPS' in self._text[i]:
				PRG_STEPS = cp(i)
				break
		self._headtext = self._text[0:HEADEND+1]
		self._statustext = self._text[HEADEND+1:STATUSEND+1]
		self._text = self._text[PRG_STEPS+1:]
	def split_line(self,line):
		L=((self._text[line]).replace(';',' ')).split()
		LINIA={'pos':line}
		LINIA['command'] = L[0]; LINIA['params'] = L[1:]
		return LINIA
	def recognize_obj(self,NR_BLOKU):
		#rozpozna najbliższy obiekt od numeru bloku
		i=NR_BLOKU; BLOK=self.split_line(i)
		OBJECT = None
		if BLOK['command'] == '2019':
			#punkt pomiarowy
			OBJECT = point_prg([BLOK])
			#na razie rozpoznaje tylko punkty
		return OBJECT, i
		
		

class point:
	def __init__(self,BLOKI):
		self._blocks = BLOKI
	def P(self):
		#zwraca współrzędne punktu:
		return [float(self._blocks[len(self._blocks)-1]['params'][1]), float(self._blocks[len(self._blocks)-1]['params'][2]), float(self._blocks[len(self._blocks)-1]['params'][3])]
	def meas_points(self):
		#zwraca listę z punktami pomiarowymi:
		POINTS = []
		for i in range(0,len(self._blocks)):
			if self._blocks[i]['command']=='GOTO':
				POINTS.append([float(self._blocks[i]['params'][0]), float(self._blocks[i]['params'][1]), float(self._blocks[i]['params'][2])])
		return POINTS
	def type(self):
		return self._blocks[0]['params'][0]

class point_prg(point):
	#punkt w formacie flexa
	def P(self):
		return [float(self._blocks[0]['params'][2]), float(self._blocks[0]['params'][3]), float(self._blocks[0]['params'][4])]
	def meas_points(self):
		return [self.P()]
	def type(self):
		return 'POINT'
	

class circle(point):
	def __init__(self,BLOKI):
		self._blocks = BLOKI
	def d(self):
		#wyciąga średnice z bloków:
		return float(self._blocks[len(self._blocks)-1]['params'][4])
	def O(self):
		#zwraca współrzędne punktu środkowego łuku:
		return [float(self._blocks[len(self._blocks)-1]['params'][1]), float(self._blocks[len(self._blocks)-1]['params'][2]), float(self._blocks[len(self._blocks)-1]['params'][3])]
		
class line(point):
	def __init__(self,BLOKI):
		self._blocks = BLOKI
	def P1(self):
		return [float(self._blocks[len(self._blocks)-1]['params'][1]), float(self._blocks[len(self._blocks)-1]['params'][2]), float(self._blocks[len(self._blocks)-1]['params'][3])]
	def P2(self):
		return [float(self._blocks[len(self._blocks)-1]['params'][4]), float(self._blocks[len(self._blocks)-1]['params'][5]), float(self._blocks[len(self._blocks)-1]['params'][6])]
	def a(self):
		#zwraca współczynnik a prostej
		P1=self.P1(); P2=self.P2()
		return (P1[0]-P2[0])/(P1[1]-P2[1])
	def b(self):
		P2 = self.P2()
		return P2[1]-self.a()*P2[0]
class ellipse(point):
	def __init__(self,BLOKI):
		self._blocks = BLOKI
	def F1(self):
		#zwraca współrzędne pierwszego ogniska elipsy
		return [float(self._blocks[len(self._blocks)-1]['params'][1]), float(self._blocks[len(self._blocks)-1]['params'][2]), float(self._blocks[len(self._blocks)-1]['params'][3])]
	def F2(self):
		#zwraca współrzędne drugiego ogniska elipsy
		return [float(self._blocks[len(self._blocks)-1]['params'][4]), float(self._blocks[len(self._blocks)-1]['params'][5]), float(self._blocks[len(self._blocks)-1]['params'][6])]
	def a(self):
		#zwraca długość większej półosi elipsy:
		return float(self._blocks[len(self._blocks)-1]['params'][7])
	def c(self):
		#zwraca c elipsy:
		F1 = self.F1(); F2 = self.F2()
		return (((F1[0]-F2[0])**2 + (F1[1]-F2[1])**2)**0.5)/2.
	def b(self):
		(self.a()**2-self.c()**2)**0.5
	def al(self):
		F1 = self.F1(); F2 = self.F2()
		return (F1[1]-F2[1])/(F1[0]-F2[0])
	def bl(self):
		F2 = self.F2()
		return F2[1]-self.al()*F2[0]
	
