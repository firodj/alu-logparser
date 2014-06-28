import os,sys,re
import time
from qex_0130514 import *
from collections import OrderedDict as odict

# author: Fadhil Mandaga <firodj@gmail.com>
# version: 13 mar 2014, tab
# version: 11 des 2013, date & fan
# version: 20 nov 2013
# version: 14 mei 2013

#def safe_html(val_str):
#    return val_str.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

ALUDBG = False

def parse_params(params):
	quo = False
	spc = True
	output = []
	value = ''
	for i in xrange(0, len(params)):
		if params[i] in [' ','\t']:
			if quo: value = value + params[i]
			else: spc = True 
		else:
			if params[i] == '"':
				quo = not quo
			if spc == True:
				if value != '': output.append(value)
				value = ''
			value = value + params[i]
			spc = False
	if value != '': output.append(value)
	return output

class AluConfigNode:
	def __init__(self, name_params):
		self.parent = None
		
		self.name_params = name_params
		self.name = None
		self.params = None
		
		self.child = odict()
		
	def __repr__(self):
		return 'AluConfigNode:%s' %(self.name_params)

class AluCmdError(Exception):
	def __init__(self, ele):
		self.ele = ele
	def __str__(self):
		s = self.ele.hostname
		s = s + ': ' + ' '.join(self.ele.cmd)
		return s

def fmt_today():
	now = time.gmtime()
	today="%s / %d %s %d" % (
		"Senin Selasa Rabu Kamis Jum'at Sabtu Minggu".split(' ')[now.tm_wday],
		now.tm_mday,
		"? Januari Pebruari Maret April Mei Juni Juli Agustus September Oktober Nopember Desember".split(' ')[now.tm_mon],
		now.tm_year)
	return today
			
def cmp_portid(kx, ky):
	x = re.split('[^0-9]+', kx)
	y = re.split('[^0-9]+', ky)
	for i in xrange(0, min(len(x), len(y))):
		c = cmp( int(x[i]), int(y[i]) )
		if c <> 0: return c
	c = cmp( len(x), len(y) )
	return c
	
class AluRouter:
	def __init__(self, hostname):
		self.hostname = hostname
		self.ip_loopback = None
		self.log_elements = list()
		self.chassis_info = dict()
		self.power_feeds = dict()
		self.fan_trays = dict()
		self.cf3_free = dict()
		self.system_info = dict()
		self.snmp_info = dict()
		self.cpu_idle = None
		self.ntp_server = dict()
		self.sync_if_timing = None
		self.timos_version = None
		self.bof_address = None
		self.telnet_session = dict()
		self.tacplus_auths = dict()
		self.syslog_servers = dict()
		self.snmp_traps = dict()
		
		self.valid_toreport = 0
		self.optical_ports = dict()
		self.checkslist = dict()
		self.card_details = dict()
		self.mda_details = dict()
		self.card_states = list()
		
		self.sap_using = dict()
		self.service_using = dict()
		self.service_config = odict()
		self.sap_config = odict()
		
	# AluRouter.append_log_element
	def append_log_element(self, ele):
		if ele.hostname <> self.hostname: return False
		self.log_elements.append( ele )
		if len(ele.cmd) < 1: return

		ok = None
		if ele.cmd[0] == 'show':
			if ele.cmd[1] == 'chassis': # show chassis
				ok = self.sh_chas(ele)
				if ALUDBG:
					print 'CHASSIS =',self.chassis_info
					print 'POWER =',self.power_feeds
					print 'FAN =',self.fan_trays
				if ok:
					self.valid_toreport = self.valid_toreport + 1
			elif ele.cmd[1] == 'router':
				if ele.cmd[2] == 'interface': # show router interface chassis
					ok = self.sh_rout_ifac(ele)
					if ALUDBG:
						print 'IP =',self.ip_loopback
					if ok: self.valid_toreport = self.valid_toreport + 1
				elif ele.cmd[2] in ('static-route', 'ospf', 'ldp', 'rsvp', 'bgp', 'mpls'):
					ok = self.x_show_numbers_x(ele)
			elif ele.cmd[1] == 'redundancy':
				ok = self.x_show_redundancy_synchronization(ele)
			elif ele.cmd[1] == 'service':
				if ele.cmd[2] == 'sap-using':
					ok = self.sh_svc_sapusg(ele)
				elif ele.cmd[2] == 'service-using':
					ok = self.sh_svc_svcusg(ele)
				else:
					ok = self.x_show_numbers_x(ele)
			elif ele.cmd[1] == 'system':
				if ele.cmd[2] == 'information':
					ok = self.sh_sys_info(ele)
					if ALUDBG:
						print 'SYS =',self.system_info
						print 'SNMP =',self.snmp_info
					if ok: self.valid_toreport = self.valid_toreport + 1
				elif ele.cmd[2] == 'cpu':
					ok = self.sh_sys_cpu(ele)
					if ALUDBG:
						print 'CPU Idle =', self.cpu_idle
				elif ele.cmd[2] == 'ntp':
					ok = self.sh_sys_ntp(ele)
					if ALUDBG:
						print 'NTP =',self.ntp_server
				elif ele.cmd[2] == 'sync-if-timing':
					ok = self.x_show_system_synciftiming(ele)
					if ALUDBG:
						print 'Sync-If-Timing = ',self.sync_if_timing
				elif ele.cmd[2] == 'security':
					ok = self.x_show_system_security_authentication(ele)
					if ALUDBG:
						print 'TAC+ =',self.tacplus_auths
			elif ele.cmd[1] == 'version':
				ok = self.sh_ver(ele)
				if ALUDBG:
					print 'TiMOS =',self.timos_version
			elif ele.cmd[1] == 'bof':
				ok = self.sh_bof(ele)
				if ALUDBG:
					print 'BOF =',self.bof_address
			elif ele.cmd[1] == 'port':
				if len(ele.cmd) >= 3:
					if ele.cmd[2] != 'description':
						ok = self.x_show_port_id(ele)
				else:
					ok = self.x_show_port(ele)
			elif ele.cmd[1] == 'log':
				if ele.cmd[2] == 'syslog':
					ok = self.x_show_log_syslog(ele)
					if ALUDBG:
						print 'SYSLOG =',self.syslog_servers
				elif ele.cmd[2] == 'snmp-trap-group':
					ok = self.x_show_log_snmptrapgroup(ele)
					if ALUDBG:
						print 'TRAP =',self.snmp_traps
			elif ele.cmd[1] == 'card':
				if (len(ele.cmd) > 2) and (ele.cmd[2] == 'detail'):
					ok = self.x_show_card_detail(ele)
					if ALUDBG:
						print 'CARD =',self.card_details
				elif (len(ele.cmd) > 2) and (ele.cmd[2] == 'state'):
					ok = self.x_show_card_state(ele)
					if ALUDBG:
						print 'STATE =',self.card_states
			elif ele.cmd[1] == 'mda':
				if (len(ele.cmd) > 2) and (ele.cmd[2] == 'detail'):
					ok = self.x_show_mda_detail(ele)    
		elif ele.cmd[0] == 'file':
			if ele.cmd[1] == 'dir':
				ok = self.x_file_dir(ele)
				if ALUDBG:
					print 'CF3 =',self.cf3_free

		elif ele.cmd[0] == 'admin':
			if ele.cmd[1] == 'display-config':
				ok = self.adm_disp(ele)
				if ALUDBG:
					print 'TELNET Session =',self.telnet_session

		if ok == False:
			print "AluCmdError: ", ele.hostname, ' '.join(ele.cmd)
			#raise AluCmdError(ele)
		elif ok <> None:
			self.checkslist[ ' '.join(ele.cmd) ] = ok

	def is_okay(self, cmd):
		chk = None
		if cmd in self.checkslist:
			chk = int( self.checkslist[cmd] )
		return chk

	def x_show_card_state(self, ele):
		ele.find_reset()
		
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Card State"): return False
		if not ele.find_next("^-+"): return False

		while True:
			m = ele.find_next("^([0-9A-Z\/]+)\s+(.*)", regexp_stop="^=+")
			if m == False: break
			self.card_states.append( m.group(1) )
		
		return True
			
	def x_show_card_detail(self, ele):
		ele.find_reset()
		while True:
			if not ele.find_next("^=+"): return False
			m = ele.find_next("^Card ([0-9A-Z]+)")
			if not m: break
			
			card_id = m.group(1)
			if card_id not in self.card_details:
				self.card_details[card_id] = dict()
			card_info = self.card_details[card_id]

			if not ele.find_next("^-+"): break
			# 1         iom-20g-b        iom-20g-b        up      up
			# 1                          iom-20g-b        up      up
			m = ele.find_next("^([0-9A-Z]+)\s+(.+)", regexp_stop="^=+")
			if m == False: continue
			if not m: break
			card_info['Slot'] = m.group(1).strip(' ')

			m = re.search("^([0-9a-z\-\s]+?)\s+([a-z]+)\s+([a-z]+)", m.group(2))
			if not m: continue
			
			card_types = re.split("\s+", m.group(1))
			if len(card_types) > 1:
				card_info['Provisioned Card-type'] = card_types[0]
			else:
				card_info['Provisioned Card-type'] = None
			card_info['Equipped Card-type'] = card_types[-1]
			card_info['Admin State'] = m.group(2)
			card_info['Operational State'] = m.group(3)
			
			if not ele.find_next("^Hardware Data", regexp_stop="^=+"): continue
			
			m = ele.find_next("^\s+(Part number)\s+:\s+(.*)", regexp_stop="^=+")
			if m: card_info[ m.group(1) ] = m.group(2).strip(' ')
			m = ele.find_next("^\s+(Serial number)\s+:\s+(.*)", regexp_stop="^=+")
			if m: card_info[ m.group(1) ] = m.group(2).strip(' ')
			m = ele.find_next("^\s+(Temperature)\s+:\s+(.*)", regexp_stop="^=+")
			if m: card_info[ m.group(1) ] = m.group(2).strip(' ')
			m = ele.find_next("^\s+(Current alarm state)\s+:\s+(.*)", regexp_stop="^=+")
			if m: card_info[ m.group(1) ] = m.group(2).strip(' ')

			if not ele.find_next("^-+", regexp_stop="^=+"): continue

			errors = list()
			while True:
				m = ele.find_next( "^.+Error.*", regexp_stop="^=+")
				if not m: break
				if m: errors.append( m.group(0) )
			card_info[ 'Error' ] = errors
			
		return True
		
	def x_show_mda_detail(self, ele):
		ele.find_reset()
		last_slot = ''
		while True:
			if not ele.find_next("^=+"): return False
			m = ele.find_next("^MDA ([0-9A-Z\/]+)")
			if not m: break
			
			mda_id = m.group(1)
			if mda_id not in self.mda_details:
				self.mda_details[mda_id] = dict()
			mda_info = self.mda_details[mda_id]

			if not ele.find_next("^-+"): break

			# 1     1     m10-1gb-sfp-b         m10-1gb-sfp-b         up        up
			#       1     m10-1gb-sfp-b         m10-1gb-sfp-b         up        up
			# 1     1                           m10-1gb-sfp-b         up        up
			#       1                           m10-1gb-sfp-b         up        up
			m = ele.find_next("^([0-9A-Z\s]+)([0-9]+)\s+(.*)", regexp_stop="^=+")
			if m == False: continue
			if not m: break

			cur_slot = m.group(1).strip(' ')
			if len(cur_slot) > 0: last_slot = cur_slot
			else: cur_slot = last_slot
			mda_info['Slot'] = cur_slot
			mda_info['Mda'] = m.group(2)
			
			m = re.search("^([0-9a-z\-\/\s]+?)\s+([a-z]+)\s+([a-z]+)", m.group(3))
			if not m: continue

			mda_types = re.split("\s+", m.group(1))
			if len(mda_types) > 1:
				mda_info['Provisioned Mda-type'] = mda_types[0]
			else:
				mda_info['Provisioned Mda-type'] = None           
		   
			mda_info['Equipped Mda-type'] = mda_types[-1]
			mda_info['Admin State'] = m.group(2)
			mda_info['Operational State'] = m.group(3)

			if not ele.find_next("^Hardware Data", regexp_stop="^=+"): continue
			
			m = ele.find_next("^\s+(Part number)\s+:\s+(.*)", regexp_stop="^=+")
			if m: mda_info[ m.group(1) ] = m.group(2).strip(' ')
			m = ele.find_next("^\s+(Serial number)\s+:\s+(.*)", regexp_stop="^=+")
			if m: mda_info[ m.group(1) ] = m.group(2).strip(' ')
			m = ele.find_next("^\s+(Temperature)\s+:\s+(.*)", regexp_stop="^=+")
			if m: mda_info[ m.group(1) ] = m.group(2).strip(' ')
			m = ele.find_next("^\s+(Current alarm state)\s+:\s+(.*)", regexp_stop="^=+")
			if m: mda_info[ m.group(1) ] = m.group(2).strip(' ')

			if not ele.find_next("^-+", regexp_stop="^=+"): continue

			errors = list()
			while True:
				m = ele.find_next( "^.+Error.*", regexp_stop="^=+")
				if not m: break
				if m: errors.append( m.group(0) )
			mda_info[ 'Error' ] = errors

		#print self.mda_details
		
		return True
	
	def x_show_log_snmptrapgroup(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^SNMP Trap Groups"): return False
		if not ele.find_next("^-+"): return False
		while True:
			m = ele.find_next("^([0-9]+)\s+([0-9\.\:]+)", regexp_stop="^-+")
			if not m: break
			k = m.group(2)
			v = m.group(1)
			self.snmp_traps[ k ] = v
		return True
	
	def x_show_log_syslog(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Syslog Target"): return False
		if not ele.find_next("^-+"): return False
		while True:
			m = ele.find_next("^[0-9]+\s+([0-9\.]+)\s+([0-9]+)\s+(.*)", regexp_stop="^-+")
			if not m: break
			k = m.group(1) + ':' + m.group(2)
			v = m.group(3).strip(' ')
			self.syslog_servers[ k ] = v
		return True
			
	def x_show_system_security_authentication(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Authentication"): return False
		if not ele.find_next("^-+"): return False
		while True:
			m = ele.find_next("^tacplus\s+([^\s]+)", regexp_stop="^-+")
			if not m: break
			tacplus_status = m.group(1)
			
			m = ele.find_next("^\s+([0-9\.]+)\(([0-9]+)")
			if not m: break
			tacplus_address = m.group(1) + ':' + m.group(2)
			
			self.tacplus_auths[ tacplus_address ] = tacplus_status

		m = ele.find_next("^tacplus admin [^:]+:\s+(.*)")
		if not m: return False
		
		#print m.group(1)
		return True
		
	def x_show_redundancy_synchronization(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Synchronization Information"): return False

		m = ele.find_next("^Boot\/Config Sync Status\s+:\s+(.*)")
		if not m: return False
		
		#print m.group(1)		
		return True
		
	def x_show_numbers_x(self, ele):
		ele.find_reset()
		m = ele.find_next("^=+", regexp_stop = "^MINOR: CLI .* not configured.")
		if m == False: return None
		elif not m: return False
		m = ele.find_next("^([A-Z].*)")
		if not ele.find_next("^=+"): return False

		title = m.group(0)
		if not ele.find_next("^-+"): return False
			
		match_end = [ "^(Interfaces)\s+:\s+([0-9]+)",
					  "^(Total)\s+([0-9]+)",
					  "^(No\. of [^:]+):\s+([0-9]+)",
					  "^(Total [^:]+):\s+([0-9]+)",
					  "^(Matching [^:]+):\s+([0-9]+)",
					  "^(Number of [^:]+):\s+([0-9]+)" ]
		m = ele.find_next( match_end )

		count = None
		print 'T:',title,
		if m:
			print ';A:',m.group(1), ';B:',m.group(2)
			count = m.group(2)

		return count
		
	def x_show_port(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Ports on"): return False
		return True
		
	def x_show_port_id(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^(Ethernet|SONET/SDH) Interface"): return False
		
		m = ele.find_next("^(Description\s+:\s+)(.*)")
		if m:
			indent = len(m.group(1))
			port_desc = m.group(2)
			while True:
				m = ele.find_next("^\s{%d}(.*)" % (indent,), regexp_stop="^[^\s]+")
				if not m: break
				port_desc = port_desc + " " + m.group(1)
		m = ele.find_next("^Interface\s+:\s+([^\s]+)")
		if not m: return False
		port_id = m.group(1)
		if port_id in self.optical_ports:
			port_info = self.optical_ports[port_id]
		else:
			port_info = {'Interface': port_id, 'Description': port_desc}

		re_pat = "^((Admin|Oper) (State|Status)|Configured Mode)\s+:\s+([^\s]+)\s*(.*)"
		while True:
			m = ele.find_next(re_pat,
							  regexp_stop = "^Transceiver Data")
			if m == False: break
			if not m: return False

			k = m.group(1)
			v = m.group(m.lastindex-1)
			port_info[ k ] = v
			
			rest = m.group(m.lastindex)
			if len(rest) > 0:
				
				m = re.match(re_pat, rest)
				if m:
					k = m.group(1)
					v = m.group(m.lastindex-1)
					port_info[ k ] = v
			
		if not ele.find_next("^Transceiver Data"): return False

		while True:
			m = ele.find_next("^(Transceiver Type|(Model|Serial|Part) Number|Optical .+?|Link Length .+)\s*?:\s+(.*)",
							  regexp_stop = "^=+")
			if m == False: break
			k = m.group(1)
			v = m.group(3).rstrip(" ")
			port_info[ k ] = v
			if "Link Length" in k:
				port_info[ "Link Length" ] = v.split(" ",1)[0]
	
		if not ele.find_next("^=+"): return False
		
		if ele.find_next("^Transceiver Digital Diagnostic"):
			ele.result_i = ele.result_i + 1
			start_i = ele.result_i
			
			if ele.find_next("^=+"):
				stop_i = ele.result_i-1
				port_info['Diagnostic'] = ele.get_results(start_i, stop_i)

		if port_id not in self.optical_ports:
			self.optical_ports[port_id] = port_info
			
		#print port_info
		return True
	
	def adm_disp(self, ele):
		ele.find_reset()

		m = ele.find_next("^configure")
		if not m: return False
		
		m = ele.find_next("^(\s+)system")
		if not m: return False
		spc = len(m.group(1))

		m = ele.find_next("^(\s+)login-control")
		if not m: return True
		if len(m.group(1)) <= spc: return False
		spc = len(m.group(1))
		
		m = ele.find_next("^(\s+)telnet")
		if not m: return True
		if len(m.group(1)) <= spc: return False
		spc = len(m.group(1))

		while True:
			m = ele.find_next("^(\s+)(in|out)bound-max-sessions\s+(.*)", regexp_stop="^\sexit")
			if not m: break
			if len(m.group(1)) <= spc: break
			self.telnet_session[ m.group(2) ] = m.group(3) 
		
		while True:
			m = ele.find_next("^(\s+)(apipe|vprn|vpls|epipe|mirror-dest) ([0-9]+) .*create")
			if not m: break
			if m:
				spc_1 = len(m.group(1))
				svc_type = m.group(2)
				svc_id = m.group(3)

				if svc_id not in self.service_config:
					new_node = AluConfigNode( m.group(0).strip() )
					new_node.name = svc_type
					new_node.params = svc_id
					new_node.description = None
					self.service_config[svc_id] = new_node
				
				#print m.group(0)
				m,read_list = ele.read_next("^\s{%d}exit" %(spc_1,), "^\s{,%d}[^\s]" %(spc_1,) )
				
				spc_base = spc_1 +4
				spc_now  = spc_base
				node_now = self.service_config[svc_id]
				last_node = None

				if m:
					for j in xrange(0,len(read_list)):
						m = re.search("^(\s+)([^\s]+)(.*)", read_list[j])
						#print read_list[j]
						if not m: continue

						spc_x = len(m.group(1))
						cfg_item = m.group(0).strip()
						cfg_name = m.group(2)
						cfg_params = m.group(3).strip()
						
						if spc_x > spc_now:
							spc_now  = spc_x
							node_now = last_node
						elif spc_x < spc_now:
							spc_now = spc_x
							node_now = node_now.parent
						elif spc_x == spc_base:
							if cfg_name == 'description':
								svc_desc = cfg_params.strip('" ')
								node_now.description = svc_desc
							
						if cfg_name == 'exit': continue
						
						if cfg_item not in node_now.child:
							new_node = AluConfigNode(cfg_item)
							new_node.name = cfg_name
							new_node.params = cfg_params
							new_node.parent = node_now
							node_now.child[ cfg_item ] = new_node
							
						if cfg_name == 'sap':
							x = cfg_params.split()
							self.sap_config[x[0]] = node_now.child[ cfg_item ]

						last_node = node_now.child[ cfg_item ]

		return True
	
	def sh_bof(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^BOF \(Memory\)"): return False
		m = ele.find_next("^\s+address\s+([0-9\.]+\/[0-9]+)")
		if m: self.bof_address = m.group(1)
		return True
		
	def sh_ver(self, ele):
		ele.find_reset()
		m = ele.find_next("^(TiMOS-[^\s]+)")
		if not m: return False
		self.timos_version = m.group(1)
		return True
	
	def x_show_system_synciftiming(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^System Interface Timing"): return False
		m = ele.find_next("^System Status (C[A-Z]+ [A-Z]+)\s+:\s+(.*)")
		if not m: return False
		self.sync_if_timing = [ m.group(1).strip(' '), m.group(2).strip(' ') ]
		
		m = ele.find_next("^Reference Order\s+:\s+(.*)")
		if not m: return False

		regexp_ref = ["^Reference ([^:]+)", "^(External) Reference (In[^:]+)"]
		while True:
			m = ele.find_next(regexp_ref)
			if not m: break

			ref_name = m.group(1)
			if ref_name == "External": ref_name = ref_name + ' ' + m.group(2)
			
			#print ref_name
		return True
	
	def sh_sys_ntp(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^NTP Active Associations"): return False
		if not ele.find_next("^-+"): return False
		while True:
			m = ele.find_next("^([a-z]+)\s+([0-9\.]+)", regexp_stop="^=+")
			if not m: break
			ntp_st = m.group(1)
			ntp_ip = m.group(2)
			if m:
				if ntp_st in self.ntp_server:
					self.ntp_server[ ntp_st ] = self.ntp_server[ ntp_st ] + '; ' + ntp_ip
				else:
					self.ntp_server[ ntp_st ] = ntp_ip
		return True
	
	def sh_sys_cpu(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^CPU Utilization"): return False
		m = ele.find_next("^\s*Idle\s+[0-9,]+\s+([0-9\.]+)%")
		if m: self.cpu_idle = m.group(1)
		return True
	
	def sh_sys_info(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^System Information"): return False
		while True:
			m = ele.find_next("^(System|SNMP)\s+(.+?)\s+:\s+(.*)")
			if not m: break
			if m.group(1) == "System":
				self.system_info[ m.group(2) ] = m.group(3).strip(" ")
			if m.group(1) == "SNMP":
				self.snmp_info[ m.group(2) ] = m.group(3).strip(" ")
			
		return True
		
	def x_file_dir(self, ele):
		ele.find_reset()
		m = ele.find_next("Volume.+cf3.+slot\s+(.+?)\s+")
		if not m: return True # True some times no cf
		cf3_slot = m.group(1)
		m = ele.find_next("\s+([0-9]+)\s+bytes free\.")
		if not m: return True
		self.cf3_free[ cf3_slot ] = m.group(1)
		return True
		
	def sh_rout_ifac(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Interface Table"): return False
		if not ele.find_next("^-+"): return False
		while True:
			m = ele.find_next("^system\s+", regexp_stop="^-+")
			if m == False: break
			if not m: continue
			m = ele.find_next("^\s{3}([0-9\.]+)\/", regexp_stop="^-+")
			if m: self.ip_loopback = m.group(1)
		self.x_show_numbers_x(ele)
		return True
		
	def sh_chas(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Chassis Information"): return False
		m = ele.find_next("^\s+Name\s+:\s+(.*)")
		if m: self.chassis_info['Name'] = m.group(1).rstrip(" ")
		m = ele.find_next("^\s+Type\s+:\s+(.*)")
		if m: self.chassis_info['Type'] = m.group(1).rstrip(" ")
		m = ele.find_next("^\s+Location\s+:\s+(.*)")
		if m: self.chassis_info['Location'] = m.group(1).rstrip(" ")
		
		if not ele.find_next("^(  )?Hardware Data", regexp_stop="^-+"): return False
		m = ele.find_next("^\s+Part number\s+:\s+(.*)", regexp_stop="^-+")
		if m: self.chassis_info['Part number'] = m.group(1).strip(' ')
		m = ele.find_next("^\s+Serial number\s+:\s+(.*)", regexp_stop="^-+")
		if m: self.chassis_info['Serial number'] = m.group(1).strip(' ')

		if not ele.find_next("^-+"): return False
		if not ele.find_next("^Environment Information"): return False
		m = ele.find_next("^\s+Number of fan trays\s+:\s+([^\s]*)", regexp_stop="^\s+Fan Information")
		if m == None: return False
		num_fantrays = 0
		if m:
			num_fantrays = int(m.group(1))
			regexp_fantray = "^\s+Fan tray number\s+:\s+(.*)"
			regexp_statusspeed = "\s+(Status|Speed)\s+:\s+(.*)"
			for i in xrange(num_fantrays):
				m = ele.find_next(regexp_fantray)
				if not m: break
				fantray_slot = m.group(1).strip(' ')
				fantray_status = None
				fantray_speed = None
				for j in xrange(2):
					m = ele.find_next("\s+(Status|Speed)\s+:\s+(.*)", regexp_stop=regexp_fantray)
					if not m: break
				
					v = m.group(2).strip(' ')
					if (m.group(1) == 'Status'):
						fantray_status = v
					else:
						fantray_speed = v
				
					self.fan_trays[fantray_slot] = [fantray_status,fantray_speed]
		else:
			fantray_slot = '1'
			while True:
				m = ele.find_next("\s+Status\s+:\s+(.*)", regexp_stop="^-+")
				if not m: break
				v = [ m.group(1).strip(' ') ]
				m = ele.find_next("\s+Speed\s+:\s+(.*)", regexp_stop="^-+")
				if not m: break
				v.append( m.group(1).strip(' ') )
				self.fan_trays[fantray_slot] = v
				break
						  
		if not ele.find_next("^-+"): return False
		m = ele.find_next("^Power (Feed|Supply) Information")
		if not m: return False
		powertype = m.group(1)
		if powertype == "Feed":
			regexp_power = "^\s+Input power .+\s+:\s+(.*)"
		else:
			regexp_power = "^\s+Power supply number\s+:\s+(.*)"
			
		m = ele.find_next("^\s+Number of power .+\s+:\s+([^\s]*)")
		num_powfeeds = 0
		if m: num_powfeeds = int( m.group(1) )
		for i in xrange(num_powfeeds):
			m = ele.find_next(regexp_power)
			if not m: break
			powfeed_slot = m.group(1).strip(' ')
			m = ele.find_next("^\s+Status\s+:\s+(.*)", regexp_stop=regexp_power)
			if not m: continue
			self.power_feeds[powfeed_slot] = m.group(1).strip(' ')
			
		return True

	def auto_list(self, ele, max_col=0):
		# get header
		m,read_list = ele.read_next("^-+")
		if not m: return False
		
		# auto grid
		# TODO: algorithm for 2-rows
		cols = []
		spc  = True
		for i in xrange(0, len(read_list[0])):
			if read_list[0][i] in [' ','\t']: spc = True
			else: 
				if spc == True: cols.append(i)
				spc=False
				if max_col>0 and len(cols)>=max_col: break
		
		# parse header
		col_names = []
		for i in xrange(0, len(cols)):
			a=cols[i]
			b=None
			if i+1<len(cols): b=cols[i+1]
			col_name=[]
			for j in xrange(0, len(read_list)):
				v = read_list[j][a:b].strip()
				if v != '': col_name.append(v)
			col_names.append(' '.join(col_name))

		# get data
		m,read_list = ele.read_next("^-+")
		if not m: return False
			
		# parse data
		j = 0
		data_collect = dict()
		while j < len(read_list):
			k = j + 1
			# get next-line if multiple-lines
			while k < len(read_list):
				if read_list[k][0] in [' ','\t']: k=k+1	# FIXME: use 1-column index, not use [0]
				else: break
			
			col_values = []
			for i in xrange(0, len(cols)):
				a=cols[i]
				b=None
				if i+1<len(cols): b=cols[i+1]
				col_val=[]
				for l in xrange(j, k):
					v = read_list[l][a:b].strip()
					if v != '': col_val.append(v)
				col_values.append( ' '.join(col_val))
			
			# store sap information
			nfo = dict()
			for n in xrange(0, len(col_names)):
				nfo[ col_names[n] ] = col_values[n]
			data_collect[ col_values[0] ] = nfo
			
			#next
			j = k
			
		# get totals
		match_end = [ "^(Interfaces)\s+:\s+([0-9]+)",
					  "^(Total)\s+([0-9]+)",
					  "^(No\. of [^:]+):\s+([0-9]+)",
					  "^(Total [^:]+):\s+([0-9]+)",
					  "^(Matching [^:]+):\s+([0-9]+)",
					  "^(Number of [^:]+):\s+([0-9]+)" ]
		m = ele.find_next( match_end )
		if m:
			#print m.group(1),m.group(2)
			count = m.group(2)
			
		return data_collect
		
	def sh_svc_sapusg(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Service Access Points"): return False
		if not ele.find_next("^=+"): return False
		
		self.sap_using = self.auto_list(ele)
		return True
	
	def sh_svc_svcusg(self, ele):
		ele.find_reset()
		if not ele.find_next("^=+"): return False
		if not ele.find_next("^Services"): return False
		if not ele.find_next("^=+"): return False
		
		self.service_using = self.auto_list(ele, max_col=6)
		return True
		
	def analyze_findings(self):
		findings = list()
		if 'reject' in self.ntp_server:
			findings.append( ('ntp', 'NTP reject: ' + self.ntp_server['reject'] ))

		if self.sync_if_timing <> None:
			if self.sync_if_timing[1] <> 'Master Locked':
				findings.append( ('clocking', 'Clocking Status: ' + self.sync_if_timing[1]) )
				
		for slot in self.card_states:
			ids = slot.split('/', 1)
			if len(ids) > 1:
				slot_info = self.mda_details[slot]
			else:
				slot_info = self.card_details[slot]

			if 'Temperature' in slot_info:
				temp_C = slot_info['Temperature']
			else:
				temp_C = "?C"

			if 'Operational State' in slot_info:
				opr_st = slot_info['Operational State']
			else:
				opr_st = "down???"
				
			findings.append( (slot, "Slot %s: %s, temp = %s" % (slot, opr_st, temp_C) ) )
			if 'Error' in slot_info:
				for err in slot_info['Error']:
					err = re.sub("\s{2,}", '|', err.strip(' '))
					findings.append( (slot, err) )
		return findings

	def analyze_ports(self):
		findings = list()

		port_ids = self.optical_ports.keys()
		port_ids.sort( cmp = cmp_portid)
		
		for port_id in port_ids:
			port_info = self.optical_ports[port_id]

			opr_st = None
			if 'Oper State' in port_info: opr_st = port_info['Oper State']
			if 'Oper Status' in port_info: opr_st = port_info['Oper Status']
			if opr_st != 'up':
				findings.append( (port_id, 'Port %s, %s : %s' % (port_id, port_info['Description'], opr_st) ) )
			if 'Diagnostic' in port_info:
				for m in re.finditer("^([A-Z].+?)\s{2,}([+\-]?[0-9\.]+.*?!.*)?$", port_info['Diagnostic'], re.M):
					#values = re.split("\s+", m.group(2).strip(' '))
					values = re.sub("\s+", '|', m.group(2).strip(' '))
					alarm = m.group(1).strip(' ')
					findings.append( (port_id, 'Port %s, %s : %s = %s' % (port_id, port_info['Description'], alarm , values) ) )
		return findings

	def save_port(self, filename = None):
		port_ids = self.optical_ports.keys()
		port_ids.sort( cmp = cmp_portid)
		
		if len(port_ids) == 0: return

		if not filename: filename = self.hostname
		xl = QuickExcel( filename, 'opticalport' )

		# copy the first item row's style
		start_row = 4
		styles = dict()
		for c in "ABCDEFGHIJK":
			c_XX = xl.get_cell("%s%d" % (c, start_row))
			styles[c] = c_XX.style

		# count where to start
		last_hostname = None
		last_number = 0
		start_row = int(xl.rows_key[-1])+1
		for row in xrange(start_row-1, 3, -1):
			c_aX = xl.get_cell("A" + str(row))
			#c_bX = xl.get_cell("B" + str(row))
			c_cX = xl.get_cell("C" + str(row))
			c_dX = xl.get_cell("D" + str(row))

			if not c_dX or not c_dX.val:
				start_row = row
			if c_aX and c_aX.val: last_number = int(c_aX.val)
			if c_cX and c_cX.val:
				last_hostname = xl.get_cell_string( c_cX.id )
				print last_hostname
				break
		
		i = start_row
		if last_hostname <> self.hostname:
			xl.set_cell_value("A" + str(i), last_number+1)
			xl.set_cell_value("B" + str(i), self.ip_loopback)
			xl.set_cell_value("C" + str(i), self.hostname)
		
		for port_id in port_ids:
			port_info = self.optical_ports[port_id]
			xl.set_cell_value("D" + str(i), port_info['Interface'])
			
			adm_st = None
			opr_st = None
			
			if 'Admin State' in port_info: adm_st = port_info['Admin State']
			if 'Admin Status' in port_info: adm_st = port_info['Admin Status']

			if 'Oper State' in port_info: opr_st = port_info['Oper State']
			if 'Oper Status' in port_info: opr_st = port_info['Oper Status']

			if adm_st: xl.set_cell_value("E" + str(i), adm_st)
			if opr_st: xl.set_cell_value("F" + str(i), opr_st)
			
			xl.set_cell_value("G" + str(i), port_info['Description'])
			if 'Optical Compliance' in port_info:
				port_type = port_info['Optical Compliance'].split(' ',1)[0]
				if 'Link Length' in port_info:
					port_type = port_type + ' ' + port_info['Link Length']
				xl.set_cell_value("H" + str(i), port_type)
			if 'Model Number' in port_info:
				xl.set_cell_value("I" + str(i), port_info['Model Number'])
			if 'Diagnostic' in port_info:
				xl.set_cell_value("J" + str(i), port_info['Diagnostic'], True)

			for c in "ABCDEFGHIJK":
				coord = "%s%d" % (c, i)
				c_XX = xl.get_cell(coord)
				if not c_XX: c_XX = xl.create_cell(coord)
				c_XX.style = styles[c]

			i = i + 1
			
		xl.save_sheet()
		xl.save_strings()
		xl.save_xlsx()
		
	def save_xlsx(self):
		if self.valid_toreport < 3:
			print "Invalid to report", self.hostname
			return
				
		xl = QuickExcel( self.hostname, 'check7x50;05' )
		
		checkmark = u'\u2714'
		c_ad16 = xl.get_cell_string('AD16')
		if c_ad16: checkmark = c_ad16
		
		if self.is_okay("show chassis"):
			xl.set_cell_value("AD17", checkmark)
		else:
			xl.set_cell_value("AH17", checkmark)
			
		xl.set_cell_value("I3", self.chassis_info['Name'])
		xl.set_cell_value("I4", self.chassis_info['Type'])
		xl.set_cell_value("I5", self.timos_version)
		xl.set_cell_value("I6", self.chassis_info['Serial number'])
		xl.set_cell_value("I7", self.ip_loopback)
		xl.set_cell_value("Y5", self.chassis_info['Location'])
		xl.set_cell_value("I8", fmt_today())

		if self.is_okay("show version"):
			xl.set_cell_value("AD18", checkmark)
		else:
			xl.set_cell_value("AH18", checkmark)
			
		if self.is_okay("show bof"):
			xl.set_cell_value("AD19", checkmark)
		else:
			xl.set_cell_value("AH19", checkmark)
			
		if self.bof_address:
			xl.set_cell_value("I20", self.bof_address)

		if self.is_okay("show system information"):
			xl.set_cell_value("AD21", checkmark)
		else:
			xl.set_cell_value("AH21", checkmark)
			
		xl.set_cell_value("N22", self.snmp_info['Admin State'])
		xl.set_cell_value("N23", self.snmp_info['Oper State'])
		xl.set_cell_value("N24", self.snmp_info['Index Boot Status'])
		xl.set_cell_value("N25", self.snmp_info['Sync State'])

		if self.is_okay("show redundancy synchronization"):
			xl.set_cell_value("AD26", checkmark)
		else:
			xl.set_cell_value("AH26", checkmark)

		cpu_idle = float(self.cpu_idle)
		xl.set_cell_value("Z27", cpu_idle)
		if cpu_idle >= 20.0:
			xl.set_cell_value("AD27", checkmark)
		else:
			xl.set_cell_value("AH27", checkmark)
		
		oken = 0
		if 'chosen' in self.ntp_server:
			xl.set_cell_value("K29", self.ntp_server['chosen'])
			oken = oken + 1
		if 'candidate' in self.ntp_server:
			xl.set_cell_value("K30", self.ntp_server['candidate'])
			oken = oken + 1
		if oken == 2:
			xl.set_cell_value("AD28", checkmark)
		else:
			xl.set_cell_value("AH28", checkmark)

		if "10.137.32.195:49" in self.tacplus_auths:
			xl.set_cell_value("AD31", checkmark)
		else:
			xl.set_cell_value("AH31", checkmark)
			
		if 'A' in self.cf3_free:
			xl.set_cell_value("L35", int(self.cf3_free['A']))
		if 'B' in self.cf3_free:
			xl.set_cell_value("L36", int(self.cf3_free['B']))

		if self.sync_if_timing <> None:
			if self.sync_if_timing[1] == 'Master Locked':
				xl.set_cell_value("AD39", checkmark)
			else:
				xl.set_cell_value("AH39", checkmark)

		oken = 0
		if 'in' in self.telnet_session:
			v = int(self.telnet_session['in'])
			xl.set_cell_value("O46", v)
			if v == 7: oken = oken + 1
		if 'out' in self.telnet_session:
			v = int(self.telnet_session['out'])
			xl.set_cell_value("T46", v)
			if v == 7: oken = oken + 1

		if oken == 2:
			xl.set_cell_value("AD46", checkmark)
		else:
			xl.set_cell_value("AH46", checkmark)

		if "124.195.15.240:514" in self.syslog_servers:
			xl.set_cell_value("AD49", checkmark)
		else:
			xl.set_cell_value("AH49", checkmark)


		if ("124.195.19.20:162" in self.snmp_traps) and ("124.195.19.22:162" in self.snmp_traps):
			xl.set_cell_value("AD52", checkmark)
		else:
			xl.set_cell_value("AH52", checkmark)
			
		# router checks

		if self.is_okay("show router static-route"):
			xl.set_cell_value("AD72", checkmark)
		else:
			xl.set_cell_value("AH72", checkmark)
			
		if self.is_okay("show router ospf neighbor"):
			xl.set_cell_value("AD73", checkmark)
		else:
			xl.set_cell_value("AH73", checkmark)
			
		if self.is_okay("show router ospf interface"):
			xl.set_cell_value("AD74", checkmark)
		else:
			xl.set_cell_value("AH74", checkmark)
			
		if self.is_okay("show router mpls interface"):
			xl.set_cell_value("AD75", checkmark)
		else:
			xl.set_cell_value("AH75", checkmark)
			
		if self.is_okay("show router ldp interface"):
			xl.set_cell_value("AD76", checkmark)
		else:
			xl.set_cell_value("AH76", checkmark)

		if self.is_okay("show router ldp session"):
			xl.set_cell_value("AD77", checkmark)
		else:
			xl.set_cell_value("AH77", checkmark)

		if self.is_okay("show router rsvp session"):
			xl.set_cell_value("AD78", checkmark)
		else:
			xl.set_cell_value("AH78", checkmark)
			
		if self.is_okay("show router bgp summary"):
			xl.set_cell_value("AD79", checkmark)
		else:
			xl.set_cell_value("AH79", checkmark)

		# service checks
			
		if self.is_okay("show service customer"):
			xl.set_cell_value("AD83", checkmark)
		else:
			xl.set_cell_value("AH83", checkmark)

		if self.is_okay("show service service-using"):
			xl.set_cell_value("AD84", checkmark)
		else:
			xl.set_cell_value("AH84", checkmark)

		if self.is_okay("show service sdp"):
			xl.set_cell_value("AD85", checkmark)
		else:
			xl.set_cell_value("AH85", checkmark)

		if self.is_okay("show service sdp-using"):
			xl.set_cell_value("AD86", checkmark)
		else:
			xl.set_cell_value("AH86", checkmark)

		if self.is_okay("show service sap-using"):
			xl.set_cell_value("AD87", checkmark)
		else:
			xl.set_cell_value("AH87", checkmark)

		if self.is_okay("show service fdb-mac"):
			xl.set_cell_value("AD88", checkmark)
		else:
			xl.set_cell_value("AH88", checkmark)
			
		if '1' in self.fan_trays:
			xl.set_cell_value("L94", self.fan_trays['1'][0])
			xl.set_cell_value("Q94", self.fan_trays['1'][1])

		if '2' in self.fan_trays:
			xl.set_cell_value("L95", self.fan_trays['2'][0])
			xl.set_cell_value("Q95", self.fan_trays['2'][1])

		if '3' in self.fan_trays:
			xl.set_cell_value("L96", self.fan_trays['3'][0])
			xl.set_cell_value("Q96", self.fan_trays['3'][1])
			
		for k in ['1', 'A']:
			if k in self.power_feeds:
				xl.set_cell_value("L102", self.power_feeds[k])

		for k in ['2', 'B']:
			if k in self.power_feeds:
				xl.set_cell_value("L103", self.power_feeds[k])

		# remarks
		findings = self.analyze_findings()
		findings = findings + self.analyze_ports()
		
		r = 112
		for slot,msg in findings:
			xl.set_cell_value("C"+str(r), slot)
			xl.set_cell_value("E"+str(r), msg)
			r =r +  1
		
		xl.save()
		
	def save_log_html(self):        
		filename_log = "./out/html/" + self.hostname + ".html"

		if not os.path.exists('./out/html'):
			os.makedirs('./out/html')
			
		if os.path.exists(filename_log):
			f = open(filename_log, 'r+b')
			f.seek(-30, os.SEEK_END)
			data = f.read()
			m = re.search("<\/pre>\s*?<\/body>", data)
			if m: f.seek( - (30 - (m.start() + 7)), os.SEEK_END)
			print "Appending...", self.hostname
		else:
			f = open(filename_log, 'w+b')
			f.write( "<html>\n<head>\n<title>" + safehtml( self.hostname ) + " Log</title>\n" )
			f.write( "<style>\npre {font-family: Consolas, Courier New; font-size: 8pt; }\n" )
			f.write( "div {font-weight: bold; background-color: #f0f0a0; }\n" )
			f.write( ".floatingHeader {position: fixed; top: 0; visibility: hidden; }\n" )
			f.write( "</style>\n" )
			f.write( '<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.2/jquery.min.js"></script>')
			f.write( '''
<script>
	function UpdateTableHeaders() {
	   $(".persist-area").each(function() {
	   
		   var el             = $(this),
			   offset         = el.offset(),
			   scrollTop      = $(window).scrollTop(),
			   floatingHeader = $(".floatingHeader", this)
		   
		   if ((scrollTop > offset.top) && (scrollTop < offset.top + el.height())) {
			   floatingHeader.css({
				"visibility": "visible"
			   });
		   } else {
			   floatingHeader.css({
				"visibility": "hidden"
			   });      
		   };
	   });
	}
	
	// DOM Ready      
	$(function() {
	
	   var clonedHeaderRow;
	
	   $(".persist-area").each(function() {
		   clonedHeaderRow = $(".persist-header", this);
		   clonedHeaderRow
			 .before(clonedHeaderRow.clone())
			 .css("width", clonedHeaderRow.width())
			 .addClass("floatingHeader");
			 
	   });
	   
	   $(window)
		.scroll(UpdateTableHeaders)
		.trigger("scroll");
	   
	});
</script>
  ''' )
			f.write( "</head>\n<body>" )
			print "Creating...", self.hostname

		# For each log elements
		for ele in self.log_elements:
			#anchor_id = anchor_id + 1
			
			#<a name="">
			
			if ele.sfm:
				cmd_line = ele.sfm + ":" + ele.hostname
				if ele.ctx:
					cmd_line = cmd_line + ">".join(ele.ctx)
				cmd_line = cmd_line + "# " + ele.ori_cmd
				f.write( '<pre class="persist-area">' )
				f.write( '<div class="cmd_line persist-header">' )
				f.write( safehtml( cmd_line.rstrip("\n") ) )
				f.write('\n</div>')
			else:
				f.write('<pre>')
				
			for r_line in ele.results:
				f.write( safehtml( r_line ) + '\n')

			f.write( '</pre>\n' )
				
		f.write( "</body>\n</html>" )
		f.close()

	def save_log(self):
		filename_raw = "./out/log/" + self.hostname + ".log"

		if not os.path.exists('./out/log'):
			os.makedirs('./out/log')

		raw_f = open(filename_raw, "a")
		for ele in self.log_elements:
			if ele.sfm:
				cmd_line = ele.sfm + ":" + ele.hostname
				if ele.ctx:
					cmd_line = cmd_line + ">".join(ele.ctx)
				cmd_line = cmd_line + "# " + ele.ori_cmd
				
				raw_f.write(cmd_line)
				raw_f.write("\n")

			for r_line in ele.results:
				raw_f.write(r_line + "\n")
			
		raw_f.close()
	
def split_command(cmd):
	it = re.finditer('(^|\s)("[^"]*")(\s|$)', cmd)
	safex = list()
	begin = 0
	for m in it:
		p_stop, p_next = m.span()
		#eachx = re.sub("\s{2,}", " ", strx[begin:p_stop].strip() )
		for s in re.split("\s+", cmd[begin:p_stop].strip()):
			if len(s): safex.append(s)
		safex.append( m.group(2) )
		begin = p_next
	for s in re.split("\s+", cmd[begin:].strip()):
		if len(s): safex.append(s)
	return safex

class AluLogElement:
	# AluLogElement.__init__
	def __init__(self, hostname, sfm, ctx, cmd):
		self.hostname = hostname
		self.ori_cmd = cmd
		self.sfm = sfm
		if ctx:
			self.ctx = ctx.split('>')
		else:
			self.ctx = ctx
		self.cmd = split_command(cmd)
		self.results = list()
		self.result_i = 0

	# AluLogElement.find_reset
	def find_reset(self):
		self.result_i = 0

	# AluLogElement.get_results
	def get_results(self, start_i, stop_i):
		return "\n".join( self.results[start_i:stop_i] )

	# AluLogElement.find_next
	# False: means Stop, nothing to find
	# None: not found, find again next
	def find_next(self, regexp, regexp_stop = None, reading=False):
		# normalize regexp to multiple regexp
		if type(regexp) == str:
			regexp = [ regexp ]
			
		if not regexp_stop:
			regexp_stop = []
		elif type(regexp_stop) == str:
			regexp_stop = [ regexp_stop ]
		
		if reading: read_list = []
		for i in xrange(self.result_i, len(self.results)):
			# search
			for re_find in regexp:
				m = re.search(re_find, self.results[i])
				if m:
					self.result_i = i+1
					if reading: return m, read_list
					return m
			# else
			for re_stop in regexp_stop:
				m = re.search(re_stop, self.results[i])
				if m:
					self.result_i = i
					if reading: return False, read_list
					return False
					
			if reading: read_list.append(self.results[i])

		if reading: return None, read_list
		return None
	
	def read_next(self, regexp, regexp_stop = None):
		return self.find_next( regexp, regexp_stop, True)
		
class AluLogParser:
	def __init__(self):
		self.clear()

		if not os.path.exists('./out'):
			os.makedirs('./out')
		if not os.path.exists('./out/log'):
			os.makedirs('./out/log')
		
		self.alu_routers = dict()
		self.alu_routers_key = list()
		
	def clear(self):
		self.latest_element = None

	def append_log_element(self, ele):
		if not ele.hostname: return False
		if ele.hostname not in self.alu_routers:
			self.alu_routers[ele.hostname] = AluRouter(ele.hostname)
			self.alu_routers_key.append(ele.hostname)
		alu_box = self.alu_routers[ele.hostname]
		alu_box.append_log_element(ele)
		return True
		
	def open_and_parse(self, filename):
		f = open(filename, 'r')
		logout = None
		for line in f:
			
			x = re.match("(\*?[AB]):([0-9A-Za-z\-_]+)(#|([>a-z0-9]+)#)\s+(.*)$", line)
			if (x):
				sfm = x.group(1)
				hostname = x.group(2)
				cmd = x.group(5)
				ctx = x.group(4)

				if self.latest_element:
					if not self.latest_element.hostname:
						self.latest_element.hostname = hostname
					self.append_log_element( self.latest_element )
				
				self.latest_element = AluLogElement(hostname, sfm, ctx, cmd)
				if len(self.latest_element.cmd) >= 1:
					if self.latest_element.cmd[0] == 'logout':
						logout = hostname
			else:
				if not self.latest_element:
					self.latest_element = AluLogElement("", None, None, "")
				self.latest_element.results.append( line.rstrip("\r\n") )

				if logout:
					self.append_log_element( self.latest_element )
					logout = None
					self.latest_element = AluLogElement("", None, None, "")

		f.close()

		self.append_log_element( self.latest_element )
		self.latest_element = None

	def save_to_xlsx(self):
		for hostname in self.alu_routers_key:
			alu_box = self.alu_routers[hostname]
			alu_box.save_xlsx()

	def save_to_portx(self):
		for hostname in self.alu_routers_key:
			alu_box = self.alu_routers[hostname]
			alu_box.save_port('optical')

	def save_to_logs(self):
		for hostname in self.alu_routers_key:
			alu_box = self.alu_routers[hostname]
			alu_box.save_log()
			alu_box.save_log_html()            

class AluLogHistory:
	def __init__(self, filename=None):
		if not filename:
			filename = './alulog_history.txt'
		self.filename = filename
		self.history = dict()
		if os.path.exists(self.filename):
			fh = open(self.filename, 'r')
			for line in fh:
				path_f,m = line.split("\t")
				self.history[path_f] = m
			fh.close()
	
	def append_history_files(self,path_f,m):
		fh = open(self.filename, 'a')
		fh.write(path_f + "\t" + m + "\n")
		fh.close()
		
		self.history[path_f] = m
	
	def list_log_files(self,subdir):
		files = list()
		for f in os.listdir(subdir):
			path_f = subdir + "/" + f
			
			if f.rsplit('.',1)[-1].lower() <> "log": continue
			if not os.path.isfile(path_f): continue
			
			if path_f not in self.history:
				fs = os.stat(path_f)
				mt = fs.st_mtime
				
				files.append( (path_f, mt) )
				files.sort( key=lambda x:x[1])
				
				m = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mt))
				self.append_history_files(path_f,m)
				
		return files

def list_log_files(subdir):
	files = list()
	for f in os.listdir(subdir):
		if f.rsplit('.',1)[-1].lower() <> "log": continue
		if not os.path.isfile(subdir + "/" + f): continue
		fs = os.stat(subdir + "/" + f)
		files.append( (subdir + "/" + f, fs.st_mtime) )
		files.sort( key=lambda x:x[1])
	return files

def load_history_files():
	history = dict()
	if os.path.exists('./alulog_history.txt'):
		fh = open('./alulog_history.txt', 'r')
		for line in fh:
			f,m = line.split("\t")
			history[f] = m
		fh.close()
	return history

def append_history_files(f,m):
	fh = open('./alulog_history.txt', 'a')
	fh.write(f + "\t" + m + "\n")
	fh.close()
	
if __name__ == "__main__":
	log_parser = AluLogParser()
	
	history = load_history_files()
	for f,mt in list_log_files('./logs'):
		print "Processing", f, "..."
		if f not in history:
			m = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mt))
			
			log_parser.clear()
			log_parser.open_and_parse(f)
			
			append_history_files(f,m)
		else:
			print "Skip"
	
	log_parser.save_to_logs()
	log_parser.save_to_xlsx()
	log_parser.save_to_portx()
	
	print "Done."