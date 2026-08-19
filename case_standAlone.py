#from stuff import *
#from psstools import *
import numpy
import itertools
import sys
import pandas as pd
import utils_standAlone as utils
import os
from utils_standAlone import PSSE_PATH_OS, PSSE_PATH_SYS
import numpy as np

PSSE_PATH_OS = r"C:\Program Files\PTI\PSSE35\35.6\PSSBIN"
PSSE_PATH_SYS = r"C:\Program Files\PTI\PSSE35\35.6\PSSPY311"
sys.path.append(PSSE_PATH_OS)
sys.path.append(PSSE_PATH_SYS)    
os.environ['PATH'] += ';' + PSSE_PATH_OS
import psse35
import psspy
_f, _i, _s = psspy._f, psspy._i, psspy._s

#////////////////////////////////////////////////////////////////////////////////////////////////////

class object_:

	def __init__(self, **kwargs):
		for s in kwargs:
			self.__dict__[s] = kwargs[s]

	def __str__(self):
		output = list()
		for s in sorted(self.__dict__.keys()):
			if not s.startswith('_'):
				output.append("%s = %s" % (s, self.__dict__[s]))
		return '<' + ', '.join(output) + '>'

	def copy(self):
		r = object_()
		for n, v in self.__dict__.iteritems():
			if n.startswith('_') and n.endswith('_'):
				continue
			else:
				try: r.__dict__[n] = _cpy(v)
				except: pass
		return r

	def attributes(self):
		return {s for s in self.__dict__.keys() if
			(not (s.startswith('_') or isinstance(getattr(self, s), types.FunctionType)))}

	def default(self, **kwargs):
		for name, value in kwargs.iteritems():
			if not hasattr(self, name):
				setattr(self, name, value)

def __WriteData(data, obj, k):
	assert(isinstance(obj, object_))
	for Xs, X in data:
		for i, s in enumerate(Xs):
			obj.__dict__[s.lower()] = X[i][k]
	return obj

#////////////////////////////////////////////////////////////////////////////////////////////////////

def ReadShunts(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.aswshcount(subsys, 4)
	assert(err == 0)
	ws = ('NUMBER', 'TYPE', 'MODE', 'STATUS', 'ADJMETHOD')
	err, w = psspy.aswshint(subsys, 4, ws)
	assert(err == 0)
	wc = ('ID',)
	err, x = psspy.aswshchar(subsys, 4, wc)
	assert(err == 0)
	data = [(ws, w),(wc, x)]
	for k in range(n):
		swsh = object_()
		__WriteData(data, swsh, k)
		err, swsh.binit = psspy.swsdt1(swsh.number, 'BINIT')
		assert(err in (0, 4, 5))
		output[(swsh.number,str(swsh.id).strip().upper())] = swsh
	return output

def ReadShunts2(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.aswshcount(subsys, 4)
	assert(err == 0)
	ws = ('NUMBER', 'TYPE', 'MODE', 'STATUS', 'ADJMETHOD', 'BLOCKS') + tuple("STEPSBLOCK%i" % k for k in range(1, 9))
	err, w = psspy.aswshint(subsys, 4, ws)
	assert(err == 0)
	xs = ["BSTPBLOCK%i" % k for k in range(1, 9)]
	err, x = psspy.aswshreal(subsys, 4, xs)
	data = [(ws, w), (xs, x)]
	for k in range(n):
		swsh = object_()
		__WriteData(data, swsh, k)
		err, swsh.binit = psspy.swsdt1(swsh.number, 'BINIT')
		assert(err in (0, 4, 5))
		output[swsh.number] = swsh
		swsh.nsteps = numpy.zeros(8, dtype=int)
		swsh.stepsize = numpy.zeros(8, dtype=float)
		for k in range(swsh.blocks):
			swsh.nsteps[k] = eval("swsh.stepsblock%i" % (k + 1))
			swsh.stepsize[k] = eval("swsh.bstpblock%i" % (k + 1))
		a = [numpy.arange(swsh.nsteps[k] + 1) * swsh.stepsize[k] for k in range(swsh.blocks)]
		b = {sum(x) for x in itertools.product(*a)}
		swsh.options = {float("%.2f" % x) for x in b}
	return output

def ReadFixedShunts(subsys=None):
	subsys = -1 if (subsys is None) else subsys
	output = dict()
	err, n = psspy.afxshuntcount(subsys, 4)
	assert(err == 0)
	ws = ('NUMBER', 'STATUS')
	err, w = psspy.afxshuntint(subsys, 4, ws)
	assert(err <= 0)
	zs = ('SHUNTACT', 'SHUNTNOM')
	err, z = psspy.afxshuntcplx(subsys, 4, zs)
	assert(err <= 0)
	ss = ('ID', )
	err, s = psspy.afxshuntchar(subsys, 4, ss)
	assert(err <= 0)
	data = [(ws, w), (zs, z), (ss, s)]
	for k in range(n):
		fxsh = object_()
		__WriteData(data, fxsh, k)
		fxsh.id = str(fxsh.id).strip().upper()
		fxsh.key = (fxsh.number, fxsh.id)
		output[fxsh.key] = fxsh
	return output
	
#////////////////////////////////////////////////////////////////////////////////////////////////////

def ReadPlants(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.agenbuscount(subsys, 4)
	assert(err == 0)
	ws = ('NUMBER', 'STATUS', 'IREG')
	err, w = psspy.agenbusint(subsys, 4, ws)
	assert(err == 0)
	xs = ('VSPU', 'PU', 'RMPCT', 'QMAX', 'QMIN', 'PGEN', )
	err, x = psspy.agenbusreal(subsys, 4, xs)
	assert(err == 0)
	data = [(ws, w), (xs, x), ]
	for k in range(n):
		plant = object_()
		__WriteData(data, plant, k)
		plant.key = plant.number
		output[plant.key] = plant
	return output

#////////////////////////////////////////////////////////////////////////////////////////////////////

def ReadMachines(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.amachcount(subsys, 4)
	assert(err == 0)
	ws = ('NUMBER', 'STATUS')
	err, w = psspy.amachint(subsys, 4, ws)
	assert(err == 0)
	xs = ('MBASE', 'PMAX', 'QMAX', 'PMIN', 'QMIN', 'PGEN', 'QGEN')
	err, x = psspy.amachreal(subsys, 4, xs)
	assert(err == 0)
	zs = ('ZSORCE', 'PQGEN')
	err, z = psspy.amachcplx(subsys, 4, zs)
	assert(err == 0)
	ss = ('ID', 'NAME')
	err, s = psspy.amachchar(subsys, 4, ss)
	assert(err == 0)
	data = [(ws, w), (xs, x), (zs, z), (ss, s)]
	for k in range(n):
		machine = object_()
		__WriteData(data, machine, k)
		machine.id = str(machine.id).strip().upper()
		machine.key = (machine.number, machine.id)
		output[machine.key] = machine
	return output

def ReadLoads(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.aloadcount(subsys, 4)
	assert(err == 0)
	ws = ('NUMBER', 'STATUS', 'SCALE', 'AREA', 'ZONE')
	err, w = psspy.aloadint(subsys, 4, ws)
	assert(err == 0)
	zs = ('MVANOM', 'ILNOM', 'YLNOM', 'TOTALNOM')
	err, z = psspy.aloadcplx(subsys, 4, zs)
	assert(err == 0)
	ss = ('ID', 'NAME')
	err, s = psspy.aloadchar(subsys, 4, ss)
	assert(err == 0)
	data = [(ws, w), (zs, z), (ss, s)]
	for k in range(n):
		load = object_()
		__WriteData(data, load, k)
		load.id = str(load.id).strip().upper()
		load.key = (load.number, load.id)
		output[load.key] = load
	return output

def ReadBuses(subsys=None, volt=False):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.abuscount(subsys, 2)
	assert(err == 0)
	ws = ('NUMBER', 'TYPE', 'AREA', 'ZONE', 'OWNER')
	err, w = psspy.abusint(subsys, 2, ws)
	assert(err == 0)
	ss = ('NAME', )
	err, s = psspy.abuschar(subsys, 2, ss)
	assert(err == 0)
	xs = ('BASE', 'NVLMHI', 'NVLMLO')
	if volt: xs = xs + ('PU', 'ANGLE', 'ANGLED')
	err, x = psspy.abusreal(subsys, 2, xs)
	assert(err == 0)
	data = [(ws, w), (ss, s), (xs, x)]
	for k in range(n):
		bus = object_()
		__WriteData(data, bus, k)
		output[bus.number] = bus
	return output

#////////////////////////////////////////////////////////////////////////////////////////////////////

def _TapsArray(tx):
	tapCountCheck = (tx.ntposn > 1) and (tx.ntposn <= 101)
	tapRangeCheck = (tx.rmax <= 1.6) and (tx.rmin >= 0.6) and (abs(tx.rmax - tx.rmin) > 0.001)
	if (tapCountCheck and tapRangeCheck):
		t = numpy.arange(tx.ntposn, dtype=float) / float(tx.ntposn - 1)
		tx.taps = round(tx.rmin, 5) + round(tx.rmax - tx.rmin, 5) * t
	else: tx.taps = None

def ReadThreeWinding(caseData, subsys=None):
	
	def _key(u): return tuple(sorted((u.wind1number, u.wind2number, u.wind3number))) + (u.id, )
	
	def UnusedBus(caseData):
		for k in range(999999, 0, -1):
			if k not in caseData.buses:
				return k
	
	subsys = -1 if (subsys is None) else subsys
	caseData.threeWinding = dict()
	err, n = psspy.atr3count(subsys, 1, 3, 2, 1)
	assert(err == 0)
	ws = ('WIND1NUMBER', 'WIND2NUMBER', 'WIND3NUMBER', 'STATUS')
	err, w = psspy.atr3int(subsys, 1, 3, 2, 1, ws)
	assert(err <= 0)
	ss = ('ID', 'XFRNAME', )
	err, s = psspy.atr3char(subsys, 1, 3, 2, 1, ss)
	assert(err <= 0)
	data = ((ws, w), (ss, s))
	for k in range(n):
		tx = object_(windings=dict())
		__WriteData(data, tx, k)
		tx.key = _key(tx)
		fakeBus = object_(number=UnusedBus(caseData),
			type=1, area=None, zone=None, owner=None, name='STAR',
			base=0.0, nvlmhi=0.0, nvlmlo=0.0)
		tx.starBus = fakeBus
		caseData.buses[fakeBus.number] = fakeBus
		caseData.threeWinding[tx.key] = tx
	err, n = psspy.awndcount(subsys, 1, 3, 2, 1)
	assert(err == 0)
	ws = ('WNDNUM', 'CODE', 'STATUS', 'NTPOSN', 'CW',
		'WIND1NUMBER', 'WIND2NUMBER', 'WIND3NUMBER',
		'WNDBUSNUMBER', 'OTHER1NUMBER', 'OTHER2NUMBER')
	err, w = psspy.awndint(subsys, 1, 3, 2, 1, ws)
	assert(err <= 0)
	xs = ('RATIO', 'RATIOCW', 'ANGLE', 'RMAX', 'RMAXCW', 'RMIN', 'RMINCW')
	err, x = psspy.awndreal(subsys, 1, 3, 2, 1, xs)
	assert(err <= 0)
	ss = ('ID', 'XFRNAME', )
	err, s = psspy.awndchar(subsys, 1, 3, 2, 1, ss)
	assert(err <= 0)
	data = ((ws, w), (xs, x), (ss, s))
	for k in range(n):
		tw = object_()
		__WriteData(data, tw, k)
		tx = caseData.threeWinding[_key(tw)]
		tw.key = tuple(sorted((tw.wndbusnumber, tx.starBus.number))) + (tw.id, )
		tw.transformer = tx
		tx.windings[tw.key] = tw
		fake = object_(fromnumber=tw.wndbusnumber, tonumber=tx.starBus.number, id=tx.id,
			status=(1 if (tw.status is 1 and tx.status is 1) else 0),
			wind1number=tw.wndbusnumber, wind2number=tx.starBus.number,
			ratio2=1.0, ratio=tw.ratio, angle=tw.angle, cw=tw.cw,
			rmax=tw.rmax, rmin=tw.rmin, ntposn=tw.ntposn, taps=None,
			key=tw.key, threeWinding=True, tx=tx, tw=tw)
		_TapsArray(fake)
		caseData.transformers[fake.key] = fake






def ReadTransformers2(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()

	# 2winding transformers
	err, n = psspy.atrncount(subsys, 1, 3, 2, 1)
	assert(err == 0)
	ws = ('FROMNUMBER', 'TONUMBER', 'STATUS', 'METERNUMBER', 'NMETERNUMBER',
		'ICONTNUMBER', 'WIND1NUMBER', 'WIND2NUMBER', 'TABLE', 'CODE', 'NTPOSN',
		'CW', 'CZ', 'CM', 'TPSTT', 'ANSTT')
	
	err, w = psspy.atrnint(subsys, 1, 3, 2, 1, ws)
	assert(err == 0)
	xs = ('RATEA', 'RATEB', 'RATEC', 'RATIO', 'RATIO2', 'ANGLE',
		'RMAX', 'RMIN', 'VMAX', 'VMIN', 'STEP', 'CNXANG','RATIO')
	err, x = psspy.atrnreal(subsys, 1, 3, 2, 1, xs)
	assert(err == 0)
	ss = ('ID', 'XFRNAME', 'VECTORGROUP')
	err, s = psspy.atrnchar(subsys, 1, 3, 2, 1, ss)
	assert(err == 0)
	zs = ('YMAG', 'RXNOM', 'RXACT')
	err, z = psspy.atrncplx(subsys, 1, 3, 2, 1, zs)
	assert(err <= 0)
	data = [(ws, w), (xs, x), (ss, s), (zs, z)]
	for k in range(n):
		tx = object_()
		__WriteData(data, tx, k)
		tx.id = str(tx.id).strip().upper()
		tx.key = tuple(sorted((tx.wind1number, tx.wind2number))) + (str(tx.id), )
		output[tx.key] = tx
		_TapsArray(tx)
	return output









def ReadTransformers(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()

	# 2winding transformers
	err, n = psspy.atrncount(subsys, 1, 3, 2, 1)
	assert(err == 0)
	ws = ('FROMNUMBER', 'TONUMBER', 'STATUS', 'METERNUMBER', 'NMETERNUMBER',
		'ICONTNUMBER', 'WIND1NUMBER', 'WIND2NUMBER', 'TABLE', 'CODE', 'NTPOSN',
		'CW', 'CZ', 'CM', 'TPSTT', 'ANSTT')
	
	err, w = psspy.atrnint(subsys, 1, 3, 2, 1, ws)
	assert(err == 0)
	xs = ('RATEA', 'RATEB', 'RATEC', 'RATIO', 'RATIO2', 'ANGLE',
		'RMAX', 'RMIN', 'VMAX', 'VMIN', 'STEP', 'CNXANG')
	err, x = psspy.atrnreal(subsys, 1, 3, 2, 1, xs)
	assert(err == 0)
	ss = ('ID', 'XFRNAME', 'VECTORGROUP')
	err, s = psspy.atrnchar(subsys, 1, 3, 2, 1, ss)
	assert(err == 0)
	zs = ('YMAG', 'RXNOM', 'RXACT')
	err, z = psspy.atrncplx(subsys, 1, 3, 2, 1, zs)
	assert(err <= 0)
	data = [(ws, w), (xs, x), (ss, s), (zs, z)]
	for k in range(n):
		tx = object_()
		__WriteData(data, tx, k)
		tx.id = str(tx.id).strip().upper()
		tx.key = tuple(sorted((tx.wind1number, tx.wind2number))) + (tx.id, )
		output[tx.key] = tx
		_TapsArray(tx)




# 3 winding transformers
	err, n = psspy.atr3count(-1, 1, 3, 2, 1)
	assert(err == 0)
	ws = ('WIND1NUMBER', 'WIND2NUMBER', 'WIND3NUMBER', 'STATUS','NMETERNUMBER')	
	
	err, w = psspy.atr3int(subsys, 1, 3, 2, 1, ws)	
	assert(err == 0)
	xs = ('ANSTAR', 'VMSTAR')
	err, x = psspy.atr3real(subsys, 1, 3, 2, 1, xs)
	assert(err == 0)
	ss = ('ID', 'XFRNAME', 'VECTORGROUP')
	err, s = psspy.atr3char(subsys, 1, 3, 2, 1, ss)
	assert(err == 0)
	zs = ('YMAG') # we can add more if needed
	err, z = psspy.atr3cplx(subsys, 1, 3, 2, 1, zs)
	assert(err <= 0)
	data = [(ws, w), (xs, x), (ss, s)]
	for k in range(n):
		tx = object_()
		__WriteData(data, tx, k)
		tx.id = str(tx.id).strip().upper()
		tx.key = tuple(sorted((tx.wind1number, tx.wind2number, tx.wind3number))) + (tx.id, )
		output[tx.key] = tx
	return output

def ReadBranches(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.abrncount(subsys, 1, 3, 2, 1)
	assert(err == 0)
	ws = ('FROMNUMBER', 'TONUMBER', 'STATUS')
	err, w = psspy.abrnint(subsys, 1, 3, 2, 1, ws)
	assert(err == 0)
	xs = ('RATEA', 'RATEB', 'RATEC', 'LENGTH', 'CHARGING')
	err, x = psspy.abrnreal(subsys, 1, 3, 2, 1, xs)
	assert(err == 0)
	zs = ('RX', 'FROMSHNT', 'TOSHNT')
	err, z = psspy.abrncplx(subsys, 1, 3, 2, 1, zs)
	assert(err == 0)
	ss = ('ID', )
	err, s = psspy.abrnchar(subsys, 1, 3, 2, 1, ss)
	assert(err == 0)
	data = [(ws, w), (xs, x), (zs, z), (ss, s)]
	for k in range(n):
		branch = object_()
		__WriteData(data, branch, k)
		branch.id = str(branch.id).strip().upper()
		branch.key = tuple(sorted((branch.fromnumber, branch.tonumber))) + (branch.id, )
		output[branch.key] = branch
	return output

def ReadHVDCBranches(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.a2trmdccount(subsys, 3, 2)
	assert(err <= 0)
	ws = ('FROMNUMBER', 'TONUMBER')
	err, w = psspy.a2trmdcint(subsys, 3, 2, ws)
	assert(err <= 0)
	ss = ('DCNAME', )
	err, s = psspy.a2trmdcchar(subsys, 3, 2, ss)
	assert(err <= 0)
	data = [(ws, w), (ss, s)]
	for k in range(n):
		branch = object_()
		__WriteData(data, branch, k)
		branch.id = str(branch.dcname).strip().upper()
		branch.key = tuple(sorted((branch.fromnumber, branch.tonumber))) + (branch.id, )
		branch.rectifier = branch.fromnumber
		branch.inverter = branch.tonumber
		output[branch.key] = branch
	return output

#////////////////////////////////////////////////////////////////////////////////////////////////////

def ReadBusVoltages(subsys=None):
	if (subsys is None): subsys = -1
	output = dict()
	err, n = psspy.abuscount(subsys, 2)
	assert(err == 0)
	ws = ('NUMBER', 'TYPE')
	err, w = psspy.abusint(subsys, 2, ws)
	assert(err == 0)
	zs = ('VOLTAGE', )
	err, z = psspy.abuscplx(subsys, 2, zs)
	assert(err == 0)
	output = dict()
	for k in range(n):
		bus = w[0][k]
		busType = w[1][k]
		v = z[0][k]
		if busType in (1, 2, 3):
			output[bus] = v
	return output

#////////////////////////////////////////////////////////////////////////////////////////////////////

def ReadCase(caseFile=None, branches=False, shunts=False):
	if not (caseFile is None):
		err = psspy.case(caseFile)
		assert(err == 0)
	case = object_()
	case.buses = ReadBuses()
	case.machines = ReadMachines()
	case.loads = ReadLoads()
	case.baseMVA = psspy.sysmva()
	case.plants = ReadPlants()
	for k in case.machines:
		bus, id = k
		case.machines[k].bus = case.buses[bus]
	for k in case.loads:
		bus, id = k
		case.loads[k].bus = case.buses[bus]
	if branches:
		case.transformers = ReadTransformers()
		case.branches = ReadBranches()
		case.dcLines = ReadHVDCBranches()
		# ReadThreeWinding(case)
	if shunts:
		case.shunts = ReadShunts2()
		case.fixedShunts = ReadFixedShunts()
	if not (caseFile is None):
		psspy.close_powerflow()
	return case



def ReadCase2(caseFile=None, sid=-1 ):
	if not (caseFile is None):
		err = psspy.case(caseFile)
		assert(err == 0)
	case = object_()
	case.transformers = ReadTransformers2(sid)
	case.swshunts = ReadShunts(sid)
	case.fxshunts = ReadFixedShunts(sid)
	if not (caseFile is None):
		psspy.close_powerflow()
	return case


def get_case_df(case) -> pd.DataFrame:
    '''
    Iterates over case and writes all attributes to a dataframe.
    '''
    data = []
    # Static Loads
    for bus, id in case.loads:
        load = case.loads[(bus, id)]
        data.append(
            {
                'Bus': bus,
                'ID': id,
                'Name': load.name,
                'p': load.mvanom.real,
                'q': load.mvanom.imag,
                'pmax': 1000.0,
                'Status': load.status,
                'Area': int(load.area),
                'type': 'static_load'
            }
        )

    # Machines
    for bus, id in case.machines:
        mach = case.machines[(bus, id)]
        
        # Motor Load
        if mach.pmax == 0 and mach.pmin < 0:
            data.append(
                {
                    'Bus': bus,
                    'ID': id,
                    'Name': mach.name,
                    'p': mach.pgen,
                    'q': mach.qgen,
                    'pmax': mach.pmax,
                    'Status': mach.status,
                    'Area': mach.bus.area,
                    'type': 'gen'
                }
            )

        # Gen
        else:
            data.append(
                {
                    'Bus': bus,
                    'ID': id,
                    'Name': mach.name,
                    'p': mach.pgen,
                    'q': mach.qgen,
                    'pmax': mach.pmax,
                    'Status': mach.status,
                    'Area': mach.bus.area,
                    'type': 'gen'
                }
            )
    
    # Buses
    for bus_num in case.buses:
        bus = case.buses[bus_num]
        
        # Returns integer bus type code
        err, w = psspy.busint(bus_num, 'TYPE')

        # Bus Disconnected
        if (w == 4):
            status = 0
        
        # Bus Connected
        else:
            status = 1

        data.append(
            {
                'Bus': bus.number,
                'ID': None,
                'Name': bus.name,
                'p': None,
                'Status': status,
                'Area': bus.area,
                'type': 'bus',
				'Base': bus.base
            }
        )
    # Branches
    for br_id in case.branches:
        brn = case.branches[br_id]

        data.append(
            {
                'Bus': brn.fromnumber,
                'Bus2': brn.tonumber,
                'ID': brn.id,
                'Name': None,
                'p': None,
                'Status': brn.status,
                'Area': None,
                'Area2': None,
                'type': 'branch'
            }
        )	



    # trans
    for trn_id in case.transformers:
        trn = case.transformers[trn_id]
        try:
            data.append(
            {
            'Bus':sorted([trn.fromnumber, trn.tonumber], reverse=False)[0],
            'Bus2': sorted([trn.fromnumber, trn.tonumber], reverse=False)[1],
            'Bus3': 0,
            'ID': trn.id,
            'Name': None,
            'p': None,
            'Status': trn.status,
            'Area': None,
            'Area2': None,
            'type': 'trans'
            }
			)	
        except:
            data.append(
            {
            'Bus': sorted([trn.wind1number, trn.wind2number,trn.wind3number], reverse=False)[0], # this is to make sure the order of reading buses from idev and from case are the same so that we can check if the transofrmer is added to the case 
            'Bus2': sorted([trn.wind1number, trn.wind2number,trn.wind3number], reverse=False)[1],
            'Bus3': sorted([trn.wind1number, trn.wind2number,trn.wind3number], reverse=False)[2],
            'ID': trn.id,
            'Name': None,
            'p': None,
            'Status': trn.status,
            'Area': None,
            'Area2': None,
            'type': 'trans'
            }
			)		
			




	

    df = pd.DataFrame(data)
    #assign area to branches
    branch_indices = df[df['type'] == 'branch'].index

    for index in branch_indices:
        bus = df.loc[index, 'Bus']
        bus2 = df.loc[index, 'Bus2']

        area = df.loc[(df['Bus'] == bus) & (df['type'] == 'bus'), 'Area'].iloc[0]
        area2 = df.loc[(df['Bus'] == bus2) & (df['type'] == 'bus'), 'Area'].iloc[0]
        df.loc[index, 'Area'] = area
        df.loc[index, 'Area2'] = area2

    #assign area to trans
    branch_indices = df[df['type'] == 'trans'].index

    for index in branch_indices:
        bus = df.loc[index, 'Bus']
        bus2 = df.loc[index, 'Bus2']

        area = df.loc[(df['Bus'] == bus) & (df['type'] == 'bus'), 'Area'].iloc[0]
        area2 = df.loc[(df['Bus'] == bus2) & (df['type'] == 'bus'), 'Area'].iloc[0]
        df.loc[index, 'Area'] = area
        df.loc[index, 'Area2'] = area2

    # Change bus to numeric datatype
    df['Bus'] = pd.to_numeric(df['Bus'])
    df['Bus2'] = pd.to_numeric(df['Bus2'])
    df['Bus3'] = pd.to_numeric(df['Bus3'])


    df.fillna({'Bus':0, 'Bus2':0, 'Bus3':0}, inplace=True)    
    df[['Bus', 'Bus2', 'Bus3']] = df[['Bus', 'Bus2', 'Bus3']].astype('int64')
    return df

def reloadCase(workingCaseDirectory, log=None):
        ierr = psspy.close_powerflow()
        if ierr:
            # os._exit(0)
            ierr = psspy.stop_2()
            psspy.psseinit(16000) 
            _i=psspy.getdefaultint()
            _f=psspy.getdefaultreal()
            _s=psspy.getdefaultchar()
            psspy.solution_parameters_5([_i,100,_i,_i,10,_i,_i,20,0],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])			
        psspy.progress_output(1,"",[0,0])	
        psspy.report_output(1,"",[0,0])
        ierr = psspy.case(workingCaseDirectory)
        if ierr>0 and not log is None:
            log.info("Issue with loading the case.")
        psspy.report_output(2,workingCaseDirectory.replace(".sav",""),[0,0])
        psspy.progress_output(2,workingCaseDirectory.replace(".sav",".txt"),[0,0])			
	
def saveCase(output_dir, workingCase_sav, log, remove=True):
    if os.path.exists(os.path.join(output_dir,workingCase_sav )):
        workingCase_sav = os.path.join(output_dir,workingCase_sav )
    try:
        psspy.close_powerflow()
    except:
        if log!=None: log.warning("Issue with closing PSSE35 model.")

    ierr = psspy.case(workingCase_sav)
    if ierr>0:
        if log!=None:log.warning(f"Issue with opening {workingCase_sav}")
    ierr = psspy.writerawversion('34', 0, os.path.join(f"{workingCase_sav.rstrip('.sav')}.raw"))
    if ierr>0:
        if log!=None:log.warning('Issue with saving the raw file.')
        return 0
    print('psspy.writerawversion ierr=', ierr)
    ierr = psspy.writeseqversion_2('34', 0, os.path.join(f"{workingCase_sav.rstrip('.sav')}.seq"))
    if ierr>0:
        if log!=None:log.warning('Issue with saving the seq file. The saved case will have no seq data')
    try:
        psspy.close_powerflow()
    except:
        if log!=None:log.warning("Issue with closing PSSE35 model.")
    utils.downGradeSAV(output_dir, workingCase_sav, remove, log)  
    return 1



def modify_t_Tap(inputDir, idvName, tTaps):
    idv_path = os.path.join(inputDir, f"{idvName}")
    if len(tTaps)>0:
        with open(idv_path, 'r') as file:
            lines = file.readlines()
        filtered_lines = [line for line in lines if "LTAP" not in line]	
        modified_content = ''
        for tTap in tTaps:
            modified_content = modified_content+f'''BAT_LTAP, {tTap['FromBus']}, {tTap['ToBus']}, '{tTap['ID']}',  {tTap['Percent']},{tTap['tTapBus']}," ",{tTap['KV']}\n'''	
        modified_content = modified_content + ''.join(filtered_lines)
        with open(idv_path, 'w') as file:
            file.write(modified_content)    
def runIdevFromPSSE(inputDir, idvName, tTapGraph=None):
	idv_path = os.path.join(inputDir, idvName)
	if tTapGraph is not None:
		nodes = tTapGraph.search_node_by_idv(idvName)
		if len(nodes)>0:
			tTap = []
			for node in nodes:
				tTap.append(tTapGraph.get_node_value(node))
			modify_t_Tap(inputDir, idvName, tTap)
		psspy.runrspnsfile(idv_path)
	else:
		psspy.runrspnsfile(idv_path)
