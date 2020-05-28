# Erlang Library for contact center operations forecasting
# Copyright (c) 2020 by Peter Gossler
# Version 0.1.0

import math
from pathxtend.path import Path
try:
    from erlang_base import Erlang_Base
except ImportError:
    from .erlang_base import Erlang_Base
import logging

class Erlang:
    err_val_ltz = 'Value error - parameter cannot be less than 0'
    base        = Erlang_Base
    MaxLoops    = 100
    MaxAccuracy = 1E-05
    deathrate   = 0
    
    #   __init__ (float, int, int, int, int, int, bool, int, int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #   sla       - Service Level (0.85, 0.80, etc. -> sla !<= 1.00)
    #   tta       - Time to Attend to an interaction in seconds -> typically 20 seconds for voice, 3600 seconds for email
    #   ait       - Average Interaction time in seconds - the time a service representative actively works with a customer
    #   aiw       - After Interaction Wrap in seconds - the time a service representative takes to close a transaction after finishing with a customer
    #   abnt      - The time in seconds before a transaction will be abandoned
    #   max_wait  - The maximum time a customer should wait until they are connected to a service representative
    #   nv        - Controls calculations for non-voice channels (default - False = Voice)
    #   ccc       - Value of how many concurrent transactions an agent can handle
    #   interval  - The forecasting interval 15, 30, 45, 60 minutes
    #   ops_hrs   - Operational hours of the contact center e.g. 8, 16, 24
    def __init__(self, sla, tta, ait, aiw, abnt, max_wait, nv, ccc, interval, ops_hrs, avail=1.0):
        try:
            self.logger = logging.getLogger('erlang')
            self.logger.info ("ErlangC constructor called.")
            self.local_dir = str(Path.script_dir())
            self.logger = logging.getLogger()
            self.setSLA (sla)
            self.setTTA (tta)
            self.setAVAIL (avail)
            self.setABN (abnt)
            self.setMAXWAIT (max_wait)
            self.setNVCCC (nv, ccc)
            self.setInterval (interval)
            self.setOPSHRS (ops_hrs)
            self.deathrate = self.interval / self.aht
        except ValueError as ve:
                self.logger.fatal (ve)

    def print_info (self):
        myerror = "SLA: {0}% / {1} sec"
        print (myerror.format(self.sla*100, self.tta))
        myerror = "AHT: {0} sec"
        print (myerror.format(self.aht))
        myerror = "Max Wait: {0} sec"
        print (myerror.format(self.max_wait))
        myerror = "Interval: {0} min"
        print (myerror.format(self.interval/60))
        myerror = "Ops Hours: {0} hrs"
        print (myerror.format(self.ops_hrs))
        myerror = "Non-Voice: {0}"
        print (myerror.format(self.nv))
        myerror = "Concurrent Interactions: {0}"
        print (myerror.format(self.ccc))
        print ("Script Dir: " + self.local_dir)
        return

    def __DEL__ (self):
        self.logger.info ("Erlang object deleted.")

    def getSLA (self):
        return self.sla

    def setSLA (self, sla):
        if sla > 1.00 or sla <= 0:
            raise ValueError("sla: 0 < sla <= 1.00!")
        else:
            self.sla = sla
    def getTTA (self):
        return self.tta
    def setTTA (self, tta):
        if tta <= 0:
            raise ValueError("tta must be larger than 0!")
        else:
            self.tta = tta
    def getAIT (self):
        return self.ait
    def setAIT (self, ait):
        if (ait + self.aiw) <= 0:
            raise ValueError("ait + aic must be larger than 0!")
        else:
            self.aht = self.aiw + ait  / self.avail # reflect agent availability
            self.ait = int (ait)
    def getAIW (self):
        return self.aiw
    def setAIW (self, aiw):
        if (self.ait + aiw) <= 0:
            raise ValueError("ait + aic must be larger than 0!")
        else:
            self.aht = aiw + self.ait  / self.avail # reflect agent availability
            self.aiw = int (aiw)
    def getAHT (self):
        return (self.ait / self.ccc) + (self.aiw * self.ccc)      
    def getABN (self):
        return self.abnt
    def setABN (self, abnt):
        if abnt <= 0:
            raise ValueError("abnt must be larger than 0!")
        else:
            self.abnt = abnt
    def getMAXWAIT (self):
        return self.max_wait
    def setMAXWAIT (self, max_wait):
        if max_wait <= 0:
            raise ValueError("max_wait must be larger than 0!")
        else:
            self.max_wait = max_wait
    def getNV (self):
        return self.nv
    def getCCC (self):
        return self.ccc
    def setNVCCC (self, nv, ccc):
        self.nv  = nv
        if ccc <= 0:
            raise ValueError("ccc must be larger than 0!")
        else:
            if self.nv != True:
                self.ccc = 1
            else: 
                self.ccc = ccc
            self.setAHT()
    def setAHT (self):
        # Check if an agent can take more than 1 interaction, in the case of chat etc.
        if self.nv and self.ccc > 1:
            self.aht = ((self.ait / self.ccc) + self.aiw * self.ccc)
    def getInterval (self):
        return self.interval
    def setInterval (self, interval):
        if interval != 15 and interval != 30 and interval != 45 and interval != 60:
            raise ValueError("Interval must be either 15, 30, 45 or 60 minutes!")
        else:
            self.interval  = interval * 60 # convert interval from minutes to seconds
    def getOPSHRS (self):
        return self.ops_hrs
    def setOPSHRS (self, ops_hrs):
        if ops_hrs <=0 or ops_hrs > 24:  
            raise ValueError("ops_hrs must be larger than 0 hours and less than 24 hours!")
        else:
            self.ops_hrs = ops_hrs
    def getAVAIL (self):
        return self.avail
    def setAVAIL (self, avail):
        if avail < 0.1 or avail > 1.0:
                raise ValueError("Agent availability 'avail' must be larger than 0.1 and less then or equal to 1.0!")
        else: 
            self.avail = avail
            self.setAHT()
        
    ###############################################
    ### Erlang Contact Center Related Functions ###
    ###############################################

    #   -------------------------------------------------------------------------------------------
    #   ErlangB (int, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       servers    = Number of telephone lines
    #       intensity  = Arrival rate of calls / Completion rate of calls
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - Probability in % of a call being blocked.
    #   -------------------------------------------------------------------------------------------
    def ErlangB (self, servers, intensity):
        B = 0.0
        try:
            if servers < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            
            maxiterate = self.base.FixInt(servers)
            last = 1
            i = 1
            while i <= maxiterate:
                B = (intensity * last) / (i + (intensity * last))
                last = B
                i += 1
            return self.base.MinMax(B,0,1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   ErlangBExt (int, float, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       servers   = Number of telephone lines
    #       intensity = Arrival rate of calls / Completion rate of calls
    #       retry     = Number of blocked callers who will retry immediately (0.1 = 10%)
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - Probability in % of a call being blocked.
    #   -------------------------------------------------------------------------------------------
    def ErlangBExt(self, servers, intensity, retry):
        try:
            if servers < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            maxiterate = self.base.FixInt(servers)
            retries = self.base.MinMax(retry, 0, 1)
            last = 1  #for servers = 0
            i = 0
            while i < maxiterate:
                #find out the blocking factor (for servers = 2 to n)
                B = (intensity * last) / (i + (intensity * last))
                #and the increased traffic intensity
                attempts = 1 / (1 - (B * retries))
                B = (intensity * last * attempts) / (i + (intensity * last * attempts))
                last = B
                i = i + 1
            return self.base.MinMax(B, 0 ,1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   EngsetB (int, int, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       servers    = Number of telephone lines
    #       intensity  = Arrival rate of calls / Completion rate of calls
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - Probability in % of a call being blocked.
    #   -------------------------------------------------------------------------------------------
    def EngsetB (self, servers, events, intensity):
        B = 0.0
        try:
            if servers < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            maxiterate = self.base.FixInt(servers)
            last = 1
            i = 0
            while i < maxiterate:
                B = (last * (i / (events-i) * intensity)) + 1
                last = B
                i += 1
            if B == 0:
                return 0
            else:
                return self.base.MinMax(B,0,1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   ErlangC (int, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #   servers    = Number of telephone lines
    #   intensity  = Arrival rate of calls / Completion rate of calls
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - Probability in % of a transaction being placed in a queue.
    #   -------------------------------------------------------------------------------------------
    def ErlangC(self, agents, intensity):
        try:
            if agents < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            B = self.ErlangB (agents, intensity)
            C = B / (((intensity / agents) * B) + (1 - (intensity / agents)))
            return self.base.MinMax(C,0,1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   NBTrunks (float, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       intensity  = Busyhour traffic in Erlangs
    #       blocking  = blocking factor percentage e.g. 0.10  (10% of calls may receive busy tone)
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - The number of telephone lines required.
    #   -------------------------------------------------------------------------------------------
    def NBTrunks(self, intensity, blocking):
        try:
            if blocking < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            maxitrn = 65535
            sngcnt  = 0
            i       = self.base.IntCeiling(intensity)
            B       = 0
            while i < maxitrn:
                sngcnt = i
                B = self.ErlangB (sngcnt, intensity)
                if B <= blocking:
                    break
                i += 1
            if i == maxitrn:
                i = 0
            return i
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   NumberTrunks (float, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents  = Number of Agents available
    #       intensity  = Arrival rate of calls / Completion rate of calls
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - The max number of telephone lines (Trunks) required.
    #   -------------------------------------------------------------------------------------------
    def NumberTrunks(self, intensity, agents):
        try:
            if agents < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            maxitrn = 65535
            sngcnt  = 0
            i       = self.base.IntCeiling(agents)
            B       = 0
            while i < maxitrn:
                sngcnt = i
                B = self.ErlangB (sngcnt, intensity)
                if B <= 0.001:
                    break
                i += 1
            return i
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   NumberAgents (float, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       intensity = High volume traffic in Erlangs.
    #       blocking  = blocking factor percentage e.g. 0.10  (10% of calls may receive busy tone)
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) The number of agents required.
    #   -------------------------------------------------------------------------------------------
    def NumberAgents (self, intensity, blocking):
        try:
            if blocking < 0 or intensity < 0:
                raise ValueError(''.join(self.err_val_ltz))
            i    = 0
            B    = 1
            Last = 1
            while B > blocking and B > 0.001:
                i += 1
                B = (intensity * Last) / (i + (intensity * Last))
                Last = B
            return i
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   LoopingTraffic (int, float, int, float)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       trunks    = number of Trunk lines
    #       blocking  = blocking factor percentage e.g. 0.10  (10% of calls may receive busy tone)
    #       increment = traffic increase increment
    #       min_int   = Minimum traffic intensity
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) approximate blocking value.
    #   -------------------------------------------------------------------------------------------
    def LoopingTraffic (self, trunks, blocking, increment, min_int):
        try:
            minI = min_int
            B    = self.ErlangB(trunks, min_int)
            if B == blocking:
                return min_int
            incr      = increment
            intensity = min_int
            loop      = 0
            # large numbers for trunks caused locking as precision of variable intensity is reduced
            # with very high values added MaxLoop as protection
            while incr >= self.MaxAccuracy and loop < self.MaxLoops:
                B = self.ErlangB(trunks, intensity)
                if B > blocking:
                    incr = incr / 10
                    intensity = minI
                minI = intensity
                intensity = intensity + incr
                loop += 1
            return minI
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   Traffic (float, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       servers  = Number of Trunks handling the traffic
    #       blocking  = blocking factor percentage e.g. 0.10  (10% of calls may receive busy tone)
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - The max number of telephone lines (Trunks) required.
    #   -------------------------------------------------------------------------------------------
    def Traffic(self, blocking, servers):
        Trunks = 0
        try:
            Trunks = self.base.FixInt(servers)
            if blocking < 0 or servers < 1:
                raise ValueError(''.join(self.err_val_ltz))
            # find a reasonable maximum number to work with
            maxL = Trunks
            B    = self.ErlangB(servers, maxL)
            while B < blocking:
                maxL *= 2
                B = self.ErlangB(servers, maxL)
            # find the increment to start with (multiple of 10)
            incr = 1
            while incr <= (maxL / 100):
                incr *= 10
            return self.LoopingTraffic(Trunks, blocking, incr, maxL)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0
    
    ########################################################
    ### General Contact Center Metrics Related Functions ###
    ########################################################

    #   -------------------------------------------------------------------------------------------
    #   Abandon (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents      = number of available agents in the given interval
    #       transaction = the number of forecast transactions for a given interval
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht          = the average handle time.
    #       adandon_time = time in seconds before the caller will abandon.
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - Percentage of calls abandoned within interval
    #   -------------------------------------------------------------------------------------------
    def Abandon (self, agents, transactions):
        try:
            # Calculate traffic rate
            trafficrate = transactions / self.deathrate
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            C = self.ErlangC(agents, trafficrate)
            # take all queueing calls (C) and subtract calls queueing within abandontime
            A = C * math.exp((trafficrate - agents)*(self.abnt/self.aht))
            return self.base.MinMax(A,0,1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   Agents (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       service_time = target answer time in seconds e.g. 15
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       sla          = % of calls to be answered within the ServiceTime period  e.g. 0.95 (95%).
    #       aht          = the average handle time - given in the constructor of this Class.
    #       adandon_time = time in seconds before the caller will abandon .
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - Percentage of calls abandoned within interval
    #   -------------------------------------------------------------------------------------------
    def Agents (self, service_time, transactions):
        try:
            no_agents    = 0
            server       = 0
            if self.sla > 1:
                self.sla = 1
            # calculate the traffic intensity
            trafficrate = transactions / self.deathrate
            # calculate the number of Erlangs / interval
            Erlangs = self.base.FixInt((transactions * (self.aht)) / self.interval + 0.5)
            # start at number of agents for 100% utilisation
            if Erlangs < 1:
                no_agents = 1
            else:
                no_agents = int(Erlangs//1)
            utilisation = trafficrate / no_agents
            # taking a realistic approach for Utilisation less than 100%
            while utilisation >= 1:
                no_agents = no_agents + 1
                utilisation = trafficrate / no_agents
            
            maxiterate = no_agents * 100
            # try each number of agents until the correct SLA is reached
            i = 1
            while i < maxiterate:
                utilisation = trafficrate / no_agents
                if utilisation < 1:
                    server = no_agents
                    C = self.ErlangC(server, trafficrate)
                    # find the level of SLA with this number of agents
                    SLQueued = 1 - C * math.exp((trafficrate - server) * service_time / self.aht)
                    if SLQueued < 0:
                        SLQueued = 0
                    # put a limit on the accuracy required (it will never actually get to 100%)
                    if SLQueued >= self.sla:
                        i = maxiterate
                    if  SLQueued > (1 - self.MaxAccuracy):
                        i = maxiterate
                # end if
                if i != maxiterate:
                    no_agents += 1
                i += 1
            return no_agents
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0
        #   Agents
    
    #   -------------------------------------------------------------------------------------------
    #   AgentASA (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       transactions = the number of transactions received in the given interval period
    #       asa          = the Average Speed of Answer in seconds.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - Number of agents required per interval to meet ASA
    #   -------------------------------------------------------------------------------------------
    def AgentASA (self, asa, transactions):
        no_agents = 0
        server    = 0
        try:
            if asa < 0:
                self.asa = 1
            # calculate the traffic intensity
            trafficrate = transactions/self.deathrate
            # calculate the number of Erlangs/Interval
            erlangs = self.base.FixInt((transactions * self.aht) / self.interval + 0.5)
            # start calculation with 100% agent utilisation
            if erlangs < 1:
                no_agents = 1
            else:
                no_agents = int(erlangs//1)
            utilisation = trafficrate / no_agents
            # start calculation with < 100% agent utilisation
            while utilisation >= 1:
                no_agents += 1
                utilisation = trafficrate / no_agents
            maxiterate = no_agents * 100
            # try each number of agents until the correct ASA is reached
            i = 1
            while i < maxiterate:
                server = no_agents
                utilisation = trafficrate / no_agents
                C = self.ErlangC(server, trafficrate)
                answertime = C / (server * self.deathrate * (1 - utilisation))
                if (answertime * self.interval) < self.asa:
                    break
                else:
                    i += 1
                    no_agents += 1
            # end while
            return no_agents
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   ASA (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       transactions = the number of transactions received in the given interval period
    #       agents       = number of agents available 
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - Number of agents required per interval to meet ASA
    #   -------------------------------------------------------------------------------------------
    def ASA (self, agents, transactions):
        try:
            # calculate the traffic intensity
            trafficrate = transactions/self.deathrate
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            C = self.ErlangC(agents, trafficrate)
            at = C / (agents * self.deathrate * (1 - utilisation))
            return self.base.hours_to_secs(at)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   CallCapacity (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       service_time = target answer time in seconds e.g. 15
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       sla          = % of calls to be answered within the ServiceTime period  e.g. 0.95 (95%).
    #       aht          = the average handle time.
    #       adandon_time = time in seconds before the caller will abandon.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - Number of calls that can be handled by a given number of agents
    #   -------------------------------------------------------------------------------------------
    def CallCapacity(self, agents, service_time):
        try:
            # Make sure Number of agents available is a whole number
            xNoAgent = self.base.FixInt (agents)
            # Maximum number of calls at 100% utilisation
            calls = self.base.IntCeiling(self.interval / self.aht) * xNoAgent
            xAgent = self.Agents (service_time, calls)
            # Now count down call load until the current level of agents is met
            while xAgent > xNoAgent and calls > 0:
                calls -= 1
                xAgent = self.Agents(service_time, calls)
            #end while
            return calls
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   FractionalAgents (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       service_time = target answer time in seconds e.g. 15
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       sla      = % of calls to be answered within the ServiceTime period  e.g. 0.95 (95%).
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - the number of calls which can be handled by the given number of agents 
    #                   whilst still achieving the grade of service.
    #   -------------------------------------------------------------------------------------------
    def FractionalAgents(self, service_time, transactions):
        try:
            no_agents   = 0
            utilisation = 0

            if self.sla > 1:
                sla = 1
            else:
                sla = self.sla

            # calculate the traffic intensity
            trafficrate = transactions / self.deathrate
            # calculate the number of Erlangs/interval
            erlangs = self.base.FixInt((transactions * self.aht) / self.interval + 0.5)
            # start at number of agents for 100% utilisation
            if erlangs < 1:
                no_agents = 1
            else:
                no_agents = int(erlangs//1)
                utilisation = trafficrate / no_agents
            # reduce utilisation below 100%
            while utilisation >= 1:
                no_agents += 1
                utilisation = trafficrate / no_agents
            # end while
            sl_queued   = 0
            last_slq    = 0
            servers     = 0
            max_iterate = no_agents * 100
            # try each number of agents until the correct SLA is reached
            i = 1
            while i < max_iterate:
                last_slq = sl_queued
                utilisation = trafficrate / no_agents
                if utilisation < 1:
                    servers = no_agents
                    C = self.ErlangC(servers, trafficrate)
                    # find the level of SLA with this number of agents
                    sl_queued = 1 - C * math.exp((trafficrate - servers) * service_time / self.aht)
                    if sl_queued < 0:
                        sl_queued = 0
                    if sl_queued > 1:
                        sl_queued = 1
                    # put a limit on the accuracy required (it will never actually get to 100%)
                    if sl_queued >= sla or sl_queued > (1 - self.MaxAccuracy):
                        i = max_iterate
                # end if
                if i != max_iterate:
                    no_agents += 1
            # end while
            no_agents_sng = no_agents
            # do we need to calculate a fraction?
            if sl_queued > sla:
                one_agent = sl_queued - last_slq
                fract = sla - last_slq
                no_agents_sng = (fract / one_agent) + (no_agents - 1)
            # end if
            return no_agents_sng
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   FractionalCallCapacity (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       service_time = target answer time in seconds e.g. 15
    #       agents       = number of agents available 
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       sla      = % of calls to be answered within the ServiceTime period  e.g. 0.95 (95%).
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - the number of calls which can be handled by the given number of agents 
    #                   whilst still achieving the grade of service.
    #   -------------------------------------------------------------------------------------------
    def FractionalCallCapacity(self, service_time, agents):
        try:
            xNoAgent = float (agents)
            # Maximum number of calls at 100% utilisation
            calls = self.base.IntCeiling (self.interval / self.aht * xNoAgent)
            xagent = self.FractionalAgents (service_time, calls)
            # Now count down call load until the current level of agents is met
            while (xagent > xNoAgent and calls > 0):
                calls -= 1
                xagent = self.FractionalAgents (service_time, calls)
            return calls
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   Queued (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - the percentage of calls which will queue for the given number of agents.
    #   -------------------------------------------------------------------------------------------
    def Queued(self, agents, transactions):
        try:
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            return self.base.MinMax(self.ErlangC(agents, trafficrate), 0, 1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   QueueSize (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - the average queue size for a given number of agents.
    #   -------------------------------------------------------------------------------------------
    def QueueSize(self, agents, transactions):
        try:
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            C = self.ErlangC(agents, trafficrate)            
            qsize = (utilisation * C) / (1 - utilisation)
            return self.base.FixInt(qsize + 0.5)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   QueueTime (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - the average queue time for those calls which will queue.
    #   -------------------------------------------------------------------------------------------
    def QueueTime(self, agents, transactions):
        try:
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            # calculate average in the queue time for queued calls
            qtime = 1 / (agents * self.deathrate * ( 1 - utilisation))
            return self.base.hours_to_secs(qtime)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   ServiceTime (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       sla      = % of calls to be answered within the ServiceTime period  e.g. 0.95 (95%).
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - the average waiting time in which a given percentage of the calls will be 
    #                   answered.
    #   -------------------------------------------------------------------------------------------
    def ServiceTime(self, agents, transactions):
        try:
            adjust = 0
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            C = self.ErlangC(agents, trafficrate)
            # none will be queued so return 0 seconds
            if C < (1 - self.sla):
                return 0
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            # calculate average in the queue time for queued calls
            qtime = 1 / (agents * self.deathrate * ( 1 - utilisation)) * self.interval
            stime = qtime * (1 -((1 - self.sla) / C))
            ag = self.Agents(self.base.FixInt(stime),transactions)
            if ag != agents:
                adjust = 1
            return self.base.FixInt(stime + adjust)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   SLA (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #       service_time = target answer time in seconds e.g. 15
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - the service level achieved for the given number of agents.
    #   -------------------------------------------------------------------------------------------
    def SLA(self, agents, transactions, service_time):
        try:
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            # calculate average in the queue time for queued calls
            C = self.ErlangC(agents, trafficrate) 
            # now calculate SLA % as those not queuing plus those queuing
            # revised formula with thanks to Tim Bolte and Jørn Lodahl for their input
            SLQueued = 1 - C * math.exp((trafficrate - agents) * service_time / self.aht)        
            return self.base.MinMax(SLQueued, 0, 1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   Trunks (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (int) - the number of telephone lines required to service a given number of calls
    #   and agents.
    #   -------------------------------------------------------------------------------------------
    def Trunks(self, agents, transactions):
        try:
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            utilisation = trafficrate / agents
            if utilisation >= 1:
                utilisation = 0.99
            # calculate average in the queue time for queued calls
            C = self.ErlangC(agents, trafficrate) 
            answer_time = C / (agents * self.deathrate * (1 - utilisation))
            # now calculate new intensity using average life time of call 
            # (queuing time + handle time)
            R = transactions / (self.interval / (self.aht + self.base.hours_to_secs(answer_time)))
            no_trunks = self.NumberTrunks(R, agents)
            # if there is traffic (Trafficrate>0) then always return at least 1 trunk
            if no_trunks < 1 and trafficrate > 0:
                no_trunks = 1
            return no_trunks
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0

    #   -------------------------------------------------------------------------------------------
    #   Utilisation (int, int)
    #   -------------------------------------------------------------------------------------------
    #   Parameters:
    #       agents       = number of agents available 
    #       transactions = the number of transactions received in the given interval period
    #   -------------------------------------------------------------------------------------------
    #   Parameters provided in Class Constructor:
    #       aht      = the average handle time.
    #       interval = the forecasting interval 15, 30, 45, 60 minutes.
    #   -------------------------------------------------------------------------------------------
    #   Returns (float) - the utilisation percentage for the given number of agents
    #   -------------------------------------------------------------------------------------------
    def Utilisation(self, agents, transactions):
        try:
            # Calculate traffic intensity
            trafficrate = transactions / self.deathrate
            return self.base.MinMax(trafficrate / agents, 0, 1)
        except ValueError as ve:
            self.logger.error (ve)
            return 0
        except:
            self.logger.error ('General error')
            return 0
