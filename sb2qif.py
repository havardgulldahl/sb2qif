#!/usr/bin/env python
# -*- encoding:utf8 -*- #
# kate: indent-width 4;
###########################################################################
#    Copyright (C) 2006-2007 - Håvard Dahle
#    <havard@dahle.no>
#
#    Lisens: GPL2
#
# $Id: 0.7 $
#
# Dette skriptet konverterer fra kontoutskrifter lastet ned fra Skandiabanken
# til det mer generelle, de facto utvekslingsformatet QIF. QIF kan importeres
# til KMyMoney, GnuCash, Microsoft Money, Quicken, Cashbox, etc.
#
# Bruk det slik: 
# sb2qif.py [-format] <CSV-fil fra skandiabanken>,... > skandiabanken.qif
#
# eksempel:
# sb2qif.py -cashbox 9xxxxxx_200x_xxx.CSV 9xxxxxx_200x_xxx.CSV > skandiabanken_200x.qif
#
# QIF filformat:
# QIF er et rimelig skjørt og udokumentert filformat. Dette skriptet lager
# QIF-filer som fungerer fint med KMyMoney, men det kan feile i andre program.
#
# Mer om QIF:
# http://www.respmech.com/mym2qifw/qif_new.htm
# http://en.wikipedia.org/wiki/QIF
# http://svn.gnucash.org/trac/browser/gnucash/branches/1.8/src/import-export/qif-import/file-format.txt
###########################################################################


import sys, os.path, re, md5, time
from StringIO import StringIO

__doc__ = u"""
Skript som oversetter fra kontoutskrifter i CSV-format til QIF.
Håvard Gulldahl <havard@lurtgjort.no> (C) 2006-2008

Bruk:
sb2qif.py [-format] <CSV-fil fra skandiabanken>,... > skandiabanken.qif

Hvor -format er 'kmymoney' (default) eller 'cashbox' (flere formater mottas med takk).
"""

__version__ = "0.7"

class SkilleTegnFeil(Exception): pass
class QIFFeil(Exception): pass

class qifskriver:
    transaksjonstyper = {'E':[], 'I':[]}
    filkart = {}
    gammelformat = False # i 2008 (?) endret Skandiabanken cvs-formatet fra 8 til 7 kolonner

    # egenskaper som kan tilpasses for hvert eksportformat -- i utgangspunktet funker det med kmymoney
    skrivBalanse = True # Skriv "Opening balance" på toppen av qif-fila
    balanseFormat = """!Type:Bank
D%(aar)s-01-01
POpening Balance
T0.00
CX
L[konto%(konto)s]
^

"""
    skrivKategorier = True # Skriv liste over transaksjonskategorier på toppen av qif-fila
    
    datoFormat = "%(aar)s-%(mnd)s-%(dag)s" # hvordan skal datoer presenteres
    
    utgiftFormat = "-%s" # hvordan skal negative tall angis
    
    transaksjonFormat = """D%(dato)s
T%(belop)s
P%(kategori)s
N%(referanse)s
M%(tekst)s
L%(transaksjonstype)s
CR
#%(id)s
^

""" 
    
    def __init__ (self, frafiler):
        self.filer = frafiler
        for f in frafiler:
            konto, aar = self._analyser_filnavn(f)
            if not self.filkart.has_key(konto): self.filkart[konto] = {}
            if not self.filkart[konto].has_key(aar): self.filkart[konto][aar] = {'buf':None, 'inn':[]}
            self.filkart[konto][aar]['inn'].append(f)

    def konverter(self, tilfil=None):
        if tilfil is None:
            til = sys.stdout
        else:
            til = open(tilfil, "w")
        for konto in self.filkart.keys():
            for aar in self.filkart[konto].keys():
                self.filkart[konto][aar]['buf'] = StringIO()
                for f in self.filkart[konto][aar]['inn']:
                    self._konv(f, self.filkart[konto][aar]['buf'])

        if self.skrivKategorier:
            til.write(self._list_kategorier()) # list transaksjonskategorier

        if self.skrivBalanse: # skriv balanse
            for konto in self.filkart.keys():
                for aar in self.filkart[konto].keys():
                    self.filkart[konto][aar]['buf'].seek(0)
                    til.write(self.balanseFormat % locals())
                    til.write(self.filkart[konto][aar]['buf'].read())


    def _list_kategorier(self):
        s = "!Type:Cat\n"
        for innkat in self.transaksjonstyper['I']:
            s += "N%s\nI\n^\n" % innkat
        for utkat in self.transaksjonstyper['E']:
            s += "N%s\nE\n^\n" % utkat
        return s


    def konverter_ny(self):
        base, etter = os.path.splitext(self.filnavn)
        nyfil = "/tmp/" + base + ".qif"
        return self.konverter(nyfil)

    def konverter_fil(self, tilfil=None):
        if tilfil is None:
            til = sys.stdout
        #else:
            #til = open(tilfil, "w")
        self._konv(til)


    def _analyser_transaksjon(self, tr, inntekt):
        if inntekt: y = "I"
        else: y = "E"
        if self.transaksjonstyper[y].count(tr): return
        self.transaksjonstyper[y].append(tr)

    def _analyser_skilletegn(self, linje): 
        #BOKF�INGSDATO";"RENTEDATO";"BRUKSDATO";"ARKIVREFERANSE";"TYPE";"TEKST";"UT FRA KONTO";"INN P�KON
        #eller
        #"BOKF�RINGSDATO"        "RENTEDATO"     "ARKIVREFERANSE"        "TYPE"  "TEKST" "UT FRA KONTO"  "INN P� KONTO"
        for z in ('\t', ';', ','):
            p = len(linje.split(z))
            if p not in (7, 8): continue
            elif p == 7: return z
            elif p == 8:
                self.gammelformat = True
                return z
        raise SkilleTegnFeil("Kan ikke finne skilletegn")

    def _analyser_filnavn(self, filnavn):
        "Returnerer kontonnummer og årstall ut i fra filnavn"
        #97101163680_2004_nov.CSV
        #eller
        #97101163680_2008_01_01-2008_01_31.csv
        filnavn = os.path.basename(filnavn)
        base, etter = os.path.splitext(filnavn)
        mnder = ['jan', 'feb', 'mar', 'apr', 'mai', 'jun', 'jul', 'aug', 'sep', 'okt', 'nov', 'des']
        deler = filnavn.split('_')
        try:
            assert(len(deler[0]) == 11 and deler[0].isdigit()) # kontonr
            assert(len(deler[1]) == 4 and deler[0].isdigit()) # år
            return deler[0], deler[1]
        except IndexError, AssertionError:
            return "xxxxxxxxxxxx", "2001"
    
    def _konv(self, innfil, utfil, modus='csv'):
        inntegnsett = "latin1"
        uttegnsett = "utf8"
        f = file(innfil)
        topp = f.readline()
        skilletegn = self._analyser_skilletegn(topp)
        for linje in f:
            #BOKF�INGSDATO";"RENTEDATO";"BRUKSDATO";"ARKIVREFERANSE";"TYPE";"TEKST";"UT FRA KONTO";"INN P�KON
            #"2005-05-09";"2005-05-09";"09.05.2005";"93070628";"Overfrsel";"UTLBET; ID 9710022000266641";245,73;
            # eller nytt format (2008):
            #
            #"BOKF�RINGSDATO"        "RENTEDATO"     "ARKIVREFERANSE"        "TYPE"  "TEKST" "UT FRA KONTO"  "INN P� KONTO"
            # "2007-03-31"    "2007-04-01"    "90010000"      "Kreditrente"   "KREDITRENTER"          7,01

            try:
                if self.gammelformat:
                    bokdato, rentedato, bruksdato, ref, _type, tekst, ut, inn = \
                     linje.decode(inntegnsett).encode(uttegnsett).split(skilletegn)
                else:
                    bruksdato = "" # finnes ikke i nytt format
                    bokdato, rentedato, ref, _type, tekst, ut, inn = \
                     linje.decode(inntegnsett).encode(uttegnsett).split(skilletegn)
            except ValueError:
                    raise TolkeFeil(linje) ## TODO NBNBN XXXX

            d = {}
            d['aar'], d['mnd'], d['dag'] = self._strip(bokdato).split('-') # YYYY-MM-DD
            dato = self.datoFormat % d
            kategori = self._strip(tekst)
            transaksjonstype = self._strip(_type)
            #if kategori[0:4] in ('FRA-','TIL-'): # betalingsmottaker / betaler finnes i teksten
                #kategori = kategori[5:]
                #bet = re.match(r'^(.*)BETNR-\ \d+$', kategori)
                #try: kategori = bet.group(1)
                #except AttributeError: pass
            betx = re.match(r'^([A-Z ]+)?(FRA|TIL)-\ (.+)(BETNR-\ \d+)?$', kategori)
            if betx:
                kategori = betx.group(3)
                try: kategori = re.match(r'(.+) BETNR-\ \d+$', kategori).group(1)
                except AttributeError: pass

            if transaksjonstype.lower() in ('overførsel', 'overføring'): # er betalingsoverførsel
                # se om det er verdig informasjon å bruke som betalingspart
                # finn ut retning på overføringen
                if ut: transaksjonstype += " ut"
                else: transaksjonstype += " inn"
                if 'mellom egne konti' in kategori.lower(): # egen overføring
                    kategori = "Intern overføring"

            if transaksjonstype.lower().startswith("visa"):
                visax = re.match(r'^(\d{6,16})\ (.*)$', kategori) #VISAkortnummer finnes i teksten
                #kortnr = kategori[0:6]
                if visax:
                    kortnr = visax.group(1)
                    ref = self._strip(ref) + " VISA/%s" % kortnr
                    #kategori = kategori[7:]
                    kategori = visax.group(2)

            if re.match(r'^\d\d\.\d\d', kategori): #bruksdato finnes i teksten
                if d['mnd'] == "01" and kategori[3:5] == "12": # transaksjonen var forrige år
                    d['aar'] = str(int(d['aar'])-1)
                d['dag'] = kategori[0:2]
                d['mnd'] = kategori[3:5]
                #dato = self.datoFormat % locals() #(aar, mnd, dag)
                dato = self.datoFormat % d
                kategori = kategori[5:].strip()

            if transaksjonstype.lower().startswith("visa") and not ut: # penger er satt inn på visakontoen
                transaksjonstype += " innskudd"

            valx =  re.match(r'^([A-Z]{3})\ (\d+,\d\d)\ (.*)$', kategori) #Valutainfo finnes i teksten
            if valx:
                transaksjonstype += ":" + valx.group(1)
                #hva skal vi gjøre med valutabeløpet?
                #valutamengde = valx.group(2)
                kategori = valx.group(3)

            if transaksjonstype.lower().startswith("visa") and kategori[0:2] in ('S*', 'SÆ', 'M*', 'MÆ'):
                # visa-transaksjonen er varesalg (hva betyr 'Mx'?)
                transaksjonstype = transaksjonstype + "/" + kategori[0:1]
                kategori = kategori[2:]

            if re.match(r'^\d{4}\.\d{2}\.\d{5}$', kategori): #teksten er bare et kontonummer
                kategori = "FRA KONTONUMMER " + kategori

            # siste test: har noen av feltene blitt tomme?
            if len(kategori.strip()) == 0:
                kategori = "Ukjent"

            if ut: belop = self.utgiftFormat % self._penger(ut)
            else: belop = self._penger(inn.strip())
            utfil.write(self.transaksjonFormat % \
                {'dato':dato, 'belop':belop, 'kategori':kategori, 
                 'referanse':self._strip(ref), 'tekst':self._strip(tekst), 
                 'transaksjonstype':transaksjonstype, 'id':self._id(dato, tekst)})
            self._analyser_transaksjon(transaksjonstype, inntekt=bool(inn.strip()))

    def _id(self, dato, tekst):
        # lag en unik id
        try:
            import hashlib # python2.5
            _id = hashlib.md5(tekst).hexdigest()
        except ImportError:
            _id = ''.join([str(ord(z)) for z in tekst.replace(" ", "")])
        return "%s-%s-%s" % (dato, _id, time.time())

    def _strip(self, s):
        s = s.strip()
        if s[0] in ('"', "'"): s = s[1:]
        if s[-1] in ('"', "'"): s = s[:-1]
        if not s: return ""
        if s[0] == "*":
            try: s = s[1:]
            except IndexError: return ""
        s = s.replace(":", "-") # kmymoney skiller kategorier med kolon, så fjern dem

        return s

    def _penger(self, p):
        # bytt til engelsk desimaltegn etc
        mx = re.match(r'^(\d+),(\d\d)$', p)
        try: p = "%s.%s" % (mx.group(1), mx.group(2))
        except AttributeError: pass
        return p

class cashbox(qifskriver):
    """Cashboxs qif-format, dekompilert av Christopher Campbell Jensen <Christopher@xxx.net>
    
    x Ingen 'Open Balance', bare '!Type:Bank' på toppen av fila
    x Ingen liste over transaksjonskategorier
    x Datoformat på typen DD/MM/ÅÅÅÅ
    x Negative tall skrives -1.0
    x Transaksjonene har ikke kategori
    
    Denne klassen kan brukes som mal for nye formater. Subklass 'qifskriver' og sett i gang.
    """
    
    skrivBalanse = True # Skriv "Opening balance" på toppen av qif-fila
    balanseFormat = "!Type:Bank\n\n"
    skrivKategorier = False # Skriv liste over transaksjonskategorier på toppen av qif-fila
    datoFormat = "%(dag)s/%(mnd)s/%(aar)s" # hvordan skal datoer presenteres
    utgiftFormat = "-%s" # hvordan skal negative tall angis
    transaksjonFormat = """D%(dato)s
T%(belop)s
M%(tekst)s
N%(referanse)s
P%(kategori)s - %(transaksjonstype)s
#%(id)s
^

""" 


kmymoney = qifskriver ## Kmymoneys qif-forståelse er default utputt


if __name__ == '__main__':
    if not sys.argv[1:]:
        print __doc__
        sys.exit(1)
    if sys.argv[1] == '-v':
        print __version__
        sys.exit()
    if sys.argv[1][1:] in ('cashbox','kmymoney'):
        qiffer = locals()[sys.argv[1][1:]]
        sys.argv.pop(1)
    else:
        qiffer = qifskriver
    konv = qiffer(sys.argv[1:])
    konv.konverter()
