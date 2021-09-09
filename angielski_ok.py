import wx
import wx.lib.platebtn as platebtn
from gtts import gTTS
import wx.media
from wx.adv import Animation, AnimationCtrl
import os
import time
import pygame
import xlrd
import pickle

pygame.mixer.init()
# ----------------------------------------------------------------------
# wykorzystano w programnie utwory na podstawie art.29  ustawy Prawo autorskie

class StaticText(wx.StaticText):

    def SetLabel(self, label):
        if label != self.GetLabel():
            wx.StaticText.SetLabel(self, label)

class TestPanel(wx.Panel):
    def __init__(self, parent):
        # definicje zmiennych
        self.licznik = 1
        self.ktoralekcja = 1
        self.tryb = 'ANG'
        self.maxslowwlekcji = 25
        self.index = 0
        self.ilewiem = 0
        self.ilewiemprocent = 0
        self.language = 'en'
        self.cowidac = 'a'
        self.czyfilm = 'nie'
        self.czyplay = 'nie'
        self.ilelekcji = 24
        self.ktoryprzebieg = 1
        self.cb_contents = ['Witek', 'Ela']
        self.imie = 'Witek'

        wx.Panel.__init__(self, parent, -1, style=wx.TAB_TRAVERSAL | wx.CLIP_CHILDREN)

        try:
            self.mc = wx.media.MediaCtrl(self, style=wx.SIMPLE_BORDER,
                                         szBackend=wx.media.MEDIABACKEND_WMP10)

        except NotImplementedError:
            self.Destroy()
            raise

        self.timer = wx.Timer(self)
        # self.Bind(wx.EVT_TIMER, self.OnTimer)

        self.sampleList = ['one', 'two', 'three', 'four', 'five',
                           'six', 'seven', 'eight', 'nine', 'ten', 'eleven',
                           'twelve', 'thirteen', 'fourteen', 'fiveteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen',
                           'twenty', 'twenty one', 'twenty two', 'twenty three', 'twenty four']

        # lekcja
        self.nlekcja = wx.StaticText(self, -1, "lekcja 1   ", pos=(int(280 * scale), int(20 * scale)))
        self.nlekcja.SetLabel("lekcja     ")
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.nlekcja.SetBackgroundColour(wx.Colour('Blue'))
        self.nlekcja.SetFont(self.font)
        self.nlekcja.SetTransparent(255)
        # zaznaczenie ostatniego uczącego
        try:
            plik2 = '2kto.dat'
            f = pickle.load(open(plik2, 'rb'))
            self.imie = f
            self.odczytlisty()
            self.odczytzmiennych()
        except FileNotFoundError:
            self.brakpliku()

        # zerowanie lekcji
        for i in range(self.maxslowwlekcji):
            self.znam.append(0)

        #zaczyt lekcji
        fname = 'slowka.xls'
        self.xl_workbook = xlrd.open_workbook(fname)
        self.sheet_names = self.xl_workbook.sheet_names()
        self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
        self.awyr = str(self.xl_sheet.cell_value(0, 0))
        self.azdanie = str(self.xl_sheet.cell_value(0, 2))
        self.pwyr = (self.xl_sheet.cell_value(0, 1))
        self.pzdanie = (self.xl_sheet.cell_value(0, 3))

        # obrazki do sterowania
        obraz0 = wx.Image("a0.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz1 = wx.Image("a1bm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz2 = wx.Image("a2bm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz3 = wx.Image("a3bm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz4 = wx.Image("a4powbm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz5 = wx.Image("a5wykbm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz5m = wx.Image("a5mwykbm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz6 = wx.Image("a6kobm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz7 = wx.Image('a7spbm.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz8 = wx.Image('a8slbm.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap()
        obraz9 = wx.Image("a9zzbm.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()

        # menu dodatkowe
        droparrow = platebtn.PB_STYLE_DROPARROW | platebtn.PB_STYLE_SQUARE | platebtn.PB_STYLE_GRADIENT
        self.btn1 = platebtn.PlateButton(self, wx.ID_ANY, label="Uczący się", style=droparrow)
        self.btn1.SetPressColor(wx.LIGHT_GREY)
        self.menu1 = wx.Menu()
        Nowy = self.menu1.Append(wx.ID_NEW, "Tworzy nowego użytkownika")
        self.menu1.AppendSeparator()
        Wybierz = self.menu1.Append(wx.ID_ANY,"Wybierz użytkownika")
        self.menu1.AppendSeparator()
        self.menu1.AppendSeparator()
        Ang = self.menu1.Append(wx.ID_ANY, "Wybór tryb angielsko-polski")
        self.menu1.AppendSeparator()
        Pol = self.menu1.Append(wx.ID_ANY, "Wybór tryb polsko-angielski")
        self.menu1.AppendSeparator()
        self.menu1.AppendSeparator()
        Exit = self.menu1.Append(wx.ID_EXIT, "&Wyjście", u"Zakończ program")
        self.btn1.SetMenu(self.menu1)
        self.Bind(wx.EVT_MENU, self.OnExit, Exit)
        self.Bind(wx.EVT_MENU, self.OnSave, Wybierz)
        self.Bind(wx.EVT_MENU, self.OnAng, Ang)
        self.Bind(wx.EVT_MENU, self.OnPol, Pol)
        self.Bind(wx.EVT_MENU, self.OnNew, Nowy)

        # główne menu
        button1 = wx.BitmapButton(self, -1, obraz8, pos=(int(200 * scale), int(520 * scale)),
                                  size=(int(110 * scale), int(110 * scale)))
        button1.SetBitmapCurrent(obraz3)
        button1.SetBitmapPressed(obraz9)
        button1.SetToolTip('nie wiem')
        button1.SetBackgroundColour('Blue')


        button2 = wx.BitmapButton(self, -1, obraz7, pos=(int(1000 * scale), int(520 * scale)),
                                  size=(int(110 * scale), int(110 * scale)))
        button2.SetBitmapCurrent(obraz2)
        button2.SetBitmapPressed(obraz4)
        button2.SetToolTip('wiem')
        button2.SetBackgroundColour('Blue')

        button3 = wx.BitmapButton(self, -1, obraz6, pos=(int(10 * scale), int(60 * scale)),
                                  size=(int(110 * scale), int(110 * scale)))
        button3.SetBitmapPressed(obraz5)
        button3.SetToolTip('wybierz lekcję')
        button3.SetBackgroundColour('Blue')

        button0 = wx.BitmapButton(self, -1, obraz0, pos=(int(1000 * scale), int(360 * scale)),
                                  size=(int(110 * scale), int(110 * scale)))
        button0.SetBitmapPressed(obraz0)
        button0.SetToolTip('mów słowo')
        button0.SetBackgroundColour('Blue')

        # animacja
        self.anim = Animation('gzle.gif')
        self.ctrl = AnimationCtrl(self, -1, self.anim, pos=(int(100 * scale), int(200 * scale)))
        self.ctrl.Hide()
        self.anim1 = Animation('gradosc.gif')
        self.ctrl1 = AnimationCtrl(self, -1, self.anim1, pos=(int(100 * scale), int(200 * scale)))
        self.ctrl1.Hide()

        # napis FILM
        self.film = self.awyr
        if self.tryb == "ANG":
            self.button8 = wx.Button(self, -1, label=self.awyr, pos=(int(400 * scale), int(360 * scale)),
                                     size=(int(520 * scale), int(150 * scale)))

        else:
            self.button8 = wx.Button(self, -1, label=self.pwyr, pos=(int(400 * scale), int(360 * scale)),
                                     size=(int(520 * scale), int(150 * scale)))
        self.button8.SetFont(wx.Font(int(20 * scale), 74, 90, 90, False, "Calibri"))
        self.button8.SetToolTip('mów całe zdanie')
        # self film lub komentarz
        if self.tryb == "ANG":
            self.button9 = wx.StaticText(self, -1, label=self.azdanie, pos=(int(400 * scale), int(520 * scale)),
                                         size=(int(520 * scale), int(180 * scale)),
                                         style=wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_RICH2)
        else:
            self.button9 = wx.StaticText(self, -1, label=self.pzdanie, pos=(int(400 * scale), int(520 * scale)),
                                         size=(int(520 * scale), int(180 * scale)),
                                         style=wx.TE_MULTILINE | wx.TE_WORDWRAP | wx.TE_RICH2)
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_ROMAN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        self.button9.SetBackgroundColour('Yellow')
        self.button9.SetFont(self.font)
        self.button9.SetToolTip('podpowiedź')

        self.button10 = wx.Button(self, -1, label="PLAY FILM", pos=(int(880 * scale), int(200 * scale)),
                                  size=(int(100 * scale), int(100 * scale)))
        self.button10.SetToolTip('odtwórz film')
        self.button10.SetBackgroundColour('Blue')

        # lista rozwijana lekcji
        self.lb = wx.ListBox(self, -1, (int(130 * scale), int(10 * scale)), wx.Size(int(140 * scale), int(180 * scale)),
                             self.sampleList)
        self.lb.SetBackgroundColour(wx.Colour('Blue'))
        self.lb.SetForegroundColour(wx.Colour('Black'))
        self.lb.SetFont(wx.Font(int(12 * scale), wx.SWISS, wx.NORMAL, wx.NORMAL, False, 'Arial'))
        self.Bind(wx.EVT_LISTBOX, self.EvtListBox, self.lb)
        self.lb.SetSelection(0)
        self.comboBox = wx.ComboBox(self, 1, choices=self.cb_contents, pos=(int(10 * scale), int(30 * scale),),
                                    size=(int(100 * scale), int(150 * scale)), style=wx.CB_READONLY | wx.VSCROLL,
                                    value='Witek')
        self.comboBox.Hide()

        # licznik wyrazów
        self.nlicznik = wx.StaticText(self, -1, " licznik 1", pos=(int(960 * scale), int(80 * scale)))
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.nlicznik.SetFont(self.font)
        self.nlicznik.SetBackgroundColour(wx.Colour('Blue'))
        self.nlicznik.SetForegroundColour(wx.Colour('Black'))

        # ile wyrazów
        self.nmaxslow = wx.StaticText(self, -1, "ile wyrazów w lekcji :", pos=(int(960 * scale), int(40 * scale)))
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.nmaxslow.SetFont(self.font)
        self.nmaxslow.SetBackgroundColour(wx.Colour('Blue'))
        self.nmaxslow.SetForegroundColour(wx.Colour('Black'))

        # kto się uczy
        self.ktosieuczy = wx.StaticText(self, -1, "Uczy się : " + str(self.imie),
                                        pos=(int(960 * scale), int(2 * scale)))
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.ktosieuczy.SetFont(self.font)
        self.ktosieuczy.SetBackgroundColour(wx.Colour('Blue'))
        self.ktosieuczy.SetForegroundColour(wx.Colour('Black'))

        # ile wiem
        self.nwiedza = wx.StaticText(self, -1, "wiesz 0%", pos=(int(1160 * scale), int(80 * scale)))
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.nwiedza.SetFont(self.font)
        self.nwiedza.SetBackgroundColour(wx.Colour('Blue'))
        self.nwiedza.SetForegroundColour(wx.Colour('Black'))

        #tryb
        if self.tryb == "ANG":
            tryb = ' jesteś w trybie ang-pol'
        else:
            tryb= 'jesteś w trybie pol-ang'
        self.ntryb = wx.StaticText(self, -1, str(tryb), pos=(int(960 * scale), int(120 * scale)))
        self.font = wx.Font(int(20 * scale), wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.ntryb.SetFont(self.font)
        self.ntryb.SetBackgroundColour(wx.Colour('Blue'))
        self.ntryb.SetForegroundColour(wx.Colour('Black'))

        # tlo
        self.Bind(wx.EVT_ERASE_BACKGROUND, self.OnEraseBackground)

        # aktualizacja
        self.napisy()
        self.slowka()
        self.zaznaczindeks()

        # Events.
        self.Bind(wx.EVT_BUTTON, self.But1, button1)
        self.Bind(wx.EVT_BUTTON, self.But0, button0)
        self.Bind(wx.EVT_BUTTON, self.But2, button2)
        # lista rozwijana
        self.Bind(wx.EVT_BUTTON, self.OnTestButton, button3)

        # tłumaczenie dolne
        self.button9.Bind(wx.EVT_LEFT_UP, self.But9)

        # tlumaczenie gorne
        self.Bind(wx.EVT_BUTTON, self.mow, self.button8)

        # odtwarzanie filmu
        self.Bind(wx.EVT_BUTTON, self.playfilm, self.button10)
        self.Bind(wx.media.EVT_MEDIA_FINISHED, self.media_finished)

        # czas
        timer = wx.Timer(self, -1)
        self.Bind(wx.EVT_TIMER, self.OnTimer, self.timer)

    def OnExit(self, event):
        #wyjscie z programu
        self.zapislisty()
        self.zapiszmiennych()
        frame.Close()

    def OnSave(self, event):
        #wybor domyślny
        self.comboBox = wx.ComboBox(self, 1, choices=self.cb_contents, pos=(int(10 * scale), int(30 * scale),),
                                    size=(int(100 * scale), int(150 * scale)), style=wx.CB_READONLY | wx.VSCROLL,
                                    value='Witek')
        self.comboBox.SetFont(wx.Font(int(12 * scale), wx.SWISS, wx.NORMAL, wx.NORMAL, False, 'Arial'))
        self.comboBox.Bind(wx.EVT_COMBOBOX, self.OnCombo)

    def OnNew(self, event):
        # nowy uzytkownik
        self.wybor ()

    def wybor(self):
        # wprowadzenie nowego
        self.wprowadz = 'tak'
        self.wprowadzenie()
        if self.wprowadz =='tak':
            self.zakonczwprowadzanie()
        elif self.wprowadz == 'nie':
            self.comboBox.Hide()
            return
        elif self.wprowadz =='powt':
            self.wybor()

    def OnAng (self, event):
        # tryb a-p
        self.tryb = "ANG"
        self.cowidac = 'p'
        self.But9(event)
        tryb = ' jesteś w trybie ang-pol'
        self.ntryb.SetLabel(str(tryb))

    def OnPol (self,event):
        # tryb p-a
        self.tryb="POL"
        self.cowidac = 'a'
        tryb = 'jesteś w trybie pol-ang'
        self.But9(event)
        self.ntryb.SetLabel(str(tryb))

    def zakonczwprowadzanie(self):
        #nowa osoba
        if self.wprowadz == 'tak':
            self.imie=self.timie
            self.cb_contents.append(self.imie)
            if self.comboBox:
                self.comboBox.Hide()
            self.ktoralekcja = 1
            self.index = 0
            self.zerowanie()
        self.ktosieuczy.SetLabel("Uczy się :" + self.imie)
        wx.MessageBox('Uczy się : ' + self.imie)
        self.zapiszmiennych()
        self.zapislisty()
        self.odczytzmiennych()
        self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
        self.slowka()
        self.napisy()

    def wprowadzenie(self):
        #cd wprowadzania nowej osoby
        dlg = wx.TextEntryDialog(self, 'Podaj imię', 'Wprowadzanie nowej osoby', pos=(int(10 * scale), int(30 * scale)))
        if dlg.ShowModal() == wx.ID_OK:
            self.timie = dlg.GetValue()
        else:
            self.timie=''
            self.wprowadz = 'nie'
        dlg.Destroy()
        for spr in self.cb_contents:
            if spr == self.timie:
                dlg = wx.MessageDialog(None, 'To imię już jest, Poprawiasz?', u'Pytanie', wx.YES_NO | wx.ICON_QUESTION)
                if dlg.ShowModal() == wx.ID_YES:
                    dlg.Destroy()
                    self.wprowadz = 'powt'
                    return
                else:
                    dlg.Destroy()
                    self.wprowadz = 'nie'
                    return

    def OnCombo(self, event):
        #wybor osoby
        self.imie = event.GetString()
        self.comboBox.Hide()
        self.ktosieuczy.SetLabel("Uczy się :" + self.imie)
        wx.MessageBox('Uczy się : ' + event.GetEventObject().GetValue())
        self.odczytzmiennych()
        self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
        self.slowka()
        self.napisy()
        self.zaznaczindeks()

    def zapiszmiennych(self):
        #zapisuje wszyskie dane
        zapis = []
        for i in range(self.maxslowwlekcji):
            zapis.append(self.znam[i])
        zapis.append(self.licznik)
        zapis.append(self.ktoralekcja)
        zapis.append(self.index)
        zapis.append(self.ilewiem)
        zapis.append(self.ilewiemprocent)
        zapis.append(self.language)
        zapis.append(self.cowidac)
        zapis.append(self.czyfilm)
        zapis.append(self.czyplay)
        zapis.append(self.ktoryprzebieg)
        zapis.append(self.tryb)

        try:
            plik1 = '1' + str(self.imie) + '.dat'
            f = open(plik1, 'wb')
            pickle.dump(zapis, f)
            f.close()
        except FileNotFoundError:
            self.brakpliku()
        try:
            plik2 = '2kto.dat'
            f2 = open(plik2, 'wb')
            pickle.dump(self.imie, f2)
            f2.close()
        except FileNotFoundError:
            self.brakpliku()

    def odczytlisty(self):
        #odczyt ostatniej osoby
        self.cb_contents = []
        try:
            plik2 = '0kto.dat'
            f = pickle.load(open(plik2, 'rb'))
            for i in f:
                self.cb_contents.append(i)
        except FileNotFoundError:
            self.brakpliku ()

    def zapislisty(self):
        #zapis kto się uczy
        zapis2 = []
        for i in self.cb_contents:
            zapis2.append(i)
        try:
            plik2 = '0kto.dat'
            f2 = open(plik2, 'wb')
            pickle.dump(zapis2, f2)
            f2.close
        except FileNotFoundError:
            self.brakpliku ()

    def odczytzmiennych(self):
        #odczyt wszystkich danych
        self.znam = []
        try:
            plik1 = '1' + str(self.imie) + '.dat'
            f = pickle.load(open(plik1, 'rb'))
            for i in range(self.maxslowwlekcji):
                self.znam.append(f[i])
            self.licznik = f[25]
            self.ktoralekcja = f[26]
            self.index = f[27]
            self.ilewiem = f[28]
            self.ilewiemprocent = f[29]
            self.language = f[30]
            self.cowidac = f[31]
            self.czyfilm = f[32]
            self.czyplay = f[33]
            self.ktoryprzebieg = f[34]
            self.tryb = f[35]
            self.nlekcja.SetLabel("lekcja " + str((self.ktoralekcja)))
        except FileNotFoundError:
            self.brakpliku ()

    def brakpliku (self):
        #jak cos zle z plikiem
        dlg = wx.MessageDialog(None, 'Coś poszło nie tak. Wczytujemy użytkownika Witek?', u'Pytanie',
                               wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            dlg.Destroy()
            self.imie = 'Witek'
            self.ktosieuczy.SetLabel("Uczy się :" + self.imie)
            self.odczytzmiennych()
            self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
            self.slowka()
            self.napisy()
            self.zaznaczindeks()
        else:
            dlg.Destroy()
            frame.Close()

    def media_finished(self, evt):
        #zakończenie filmu
        self.czyplay = 'nie'

    def EvtListBox(self, event):
        #wybor z listy lekcji
        self.index = event.GetSelection()
        self.zaznaczindeks()
        if self.comboBox:
            self.comboBox.Hide()

    def zaznaczindeks(self):
        #zaznaczenie indeksu lekcji
        label = self.lb.GetString(self.index)
        status = 'un'
        self.lb.SetSelection(self.index)
        status2 = 'ok'

    def OnTestButton(self, evt):
        #akceptacja wyboru lekcji
        self.ktoralekcja = self.index + 1
        if self.ktoralekcja > self.ilelekcji:
            self.ktoralekcja = self.ilelekcji
        self.nlekcja.SetLabel("lekcja " + str((self.ktoralekcja)))
        self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
        self.zerowanie()
        self.slowka()

    def But0(self, event):
        # mówi słowo angielskie
        if self.comboBox:
            self.comboBox.Hide()
        self.zapisdzwiek()
        self.mc.Stop()
        filename = str('mowa/' + "s_" + self.awyr) + '.mp3'
        pygame.mixer.music.load(filename)
        pygame.mixer.music.play()

    def But1(self, event):
        # nie wiem słowa
        if self.comboBox:
            self.comboBox.Hide()
        self.ctrl.Show()
        self.animacjazlosc()
        if self.tryb == "ANG":
            self.cowidac = 'a'
        else:
            self.cowidac = 'pol'

        if self.znam[self.licznik - 1] == 10:
            self.znam[self.licznik - 1] = 6
        else:
            if self.znam[self.licznik - 1] == 6:
                self.znam[self.licznik - 1] = 6
            else:
                if self.znam[self.licznik - 1] == 2:
                    self.znam[self.licznik - 1] = 6
                else:
                    self.znam[self.licznik - 1] = 5
        self.napisy()
        self.zwiekszlicz()
        self.zapiszmiennych()

    def But2(self, event):
        #wiem słowo
        if self.comboBox:
            self.comboBox.Hide()
        self.ctrl1.Show()
        self.zapiszmiennych()
        self.animacjaradosc()
        self.wiemslowo()
        if self.tryb == "ANG":
            self.cowidac = 'a'
        else:
            self.cowidac = 'pol'
        self.zwiekszlicz()
        self.zapiszmiennych()

    def But9(self, event):
        #zmiana trybu w wyrazie -p-a / a-p
        if self.comboBox:
            self.comboBox.Hide()
        if self.cowidac == 'a':
            self.cowidac = 'p'
            self.wyswietlpol()
        else:
            self.cowidac = 'a'
            self.wyswietlang()

    def zwiekszlicz(self):
        # sprawdzenie czy już był ten wyraz
        self.licznik = self.licznik + 1
        self.sprawdzlicznik()
        self.slowka()

    def zwiekszlek(self):
        # kolejna lekcja
        if self.ilewiem == self.maxslowwlekcji:
            dlg = wx.MessageDialog(None, 'Gratulacje znasz już wszystkie słowa w tej lekcji. Chcesz nową lekcję??',
                                   u'Pytanie', wx.YES_NO | wx.ICON_QUESTION)
            if dlg.ShowModal() == wx.ID_YES:
                self.ktoralekcja += 1
                self.index +=1
                self.czywszystkielekcje()
                self.zerowanie()
                self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
                self.nlekcja.SetLabel("lekcja " + str(self.ktoralekcja))

            else:
                self.zerowanie()
        else:
            dlg = wx.MessageDialog(None, 'Powtarzamy niewyuczone słowa?', u'Pytanie', wx.YES_NO | wx.ICON_QUESTION)
            if dlg.ShowModal() == wx.ID_YES:
                self.licznik = 1
                self.ktoryprzebieg = 2
            else:
                self.ktoralekcja += 1
                self.index += 1
                self.zerowanie()
                self.czywszystkielekcje()
                self.nlekcja.SetLabel("lekcja " + str(self.ktoralekcja))

        dlg.Destroy()
        self.czywszystkielekcje()

    def czywszystkielekcje(self):
        #sprawdzenie czy wszystkie lekcje
        if self.ktoralekcja > self.ilelekcji:
            dlg = wx.MessageDialog(None, 'To już wszystkie lekcje, zaczynamy od początku?', u'Pytanie',
                                   wx.YES_NO | wx.ICON_QUESTION)
            if dlg.ShowModal() == wx.ID_YES:

                self.ktoralekcja = 1
                self.index = 0
            else:
                self.ktoralekcja -= 1
                self.index -= 1

            self.zerowanie()
            self.xl_sheet = self.xl_workbook.sheet_by_name(self.sheet_names[self.ktoralekcja - 1])
            self.nlekcja.SetLabel("lekcja " + str(self.ktoralekcja))
            dlg.Destroy()
            if self.tryb =="ANG":
                self.cowidac = 'a'
            else :
                self.cowidac ='pol'
            self.slowka()

    def mow(self, event):
        # mow zdanie
        if self.comboBox:
            self.comboBox.Hide()
        self.zapisdzwiek()
        self.mc.Stop()
        if self.cowidac == 'a':
            filename = str('mowa/' + self.awyr) + '.mp3'
        else:
            filename = str('mowa/' + self.pwyr) + '.mp3'
        pygame.mixer.music.load(filename)
        pygame.mixer.music.play()

    def zapisdzwiek(self):
        #zapis danych do wymowy
        language = 'en'
        mytext = self.awyr
        filename = str('mowa/' + "s_" + self.awyr) + '.mp3'
        if os.path.exists(filename):
            pygame.mixer.music.stop()
        else:
            myobj = gTTS(text=mytext, lang=language, slow=False)
            myobj.save(filename)

        if self.cowidac == 'a':
            language = 'en'
            mytext = self.azdanie
            filename = str('mowa/' + self.awyr) + '.mp3'
            if os.path.exists(filename):
                pygame.mixer.music.stop()
            else:
                myobj = gTTS(text=mytext, lang=language, slow=False)
                myobj.save(filename)

        else:
            language = 'pl'
            mytext = self.pzdanie
            filename = str('mowa/' + self.pwyr) + '.mp3'
            if os.path.exists(filename):
                pygame.mixer.music.stop()
            else:
                myobj = gTTS(text=mytext, lang=language, slow=False)
                myobj.save(filename)

    def slowka(self):
        # odczyt danych do wyrazu/zdania
        self.mc.Stop()
        if self.ktoryprzebieg == 2:
            self.szukajwolnego()
        self.awyr = str(self.xl_sheet.cell_value(self.licznik - 1, 0))
        self.azdanie = str(self.xl_sheet.cell_value(self.licznik - 1, 2))
        self.pwyr = (self.xl_sheet.cell_value(self.licznik - 1, 1))
        self.pzdanie = (self.xl_sheet.cell_value(self.licznik - 1, 3))
        self.czyfilm = (self.xl_sheet.cell_value(self.licznik - 1, 10))
        if self.czyfilm == 'ok':
            self.nazwafilmu = ('filmy/' + str(self.awyr) + '.mp4')
            self.mc = wx.media.MediaCtrl(self, style=wx.SIMPLE_BORDER,
                                         szBackend=wx.media.MEDIABACKEND_WMP10, pos=(int(400 * scale), int(10 * scale)),
                                         size=(int(480 * scale), int(320 * scale)))

            self.mc.Load(self.nazwafilmu)
            self.nazwafilmu = 'PLAY FILM'
            self.czyplay = 'nie'
        else:
            self.nazwafilmu = 'BRAK'
            self.mc.Hide()
        self.button10.SetLabel(str(self.nazwafilmu))
        self.napisy()
        if self.tryb =="ANG":
            self.wyswietlang()
        else:
            self.wyswietlpol()
        self.zapisdzwiek()

    def wyswietlpol(self):
        # wyświetlenie w trybie pol-ang
        self.button9.SetLabel(self.pzdanie)
        self.button8.SetLabel(self.pwyr)

    def wyswietlang(self):
        # wyświetlenie w trybie ang-pol
        self.button9.SetLabel(self.azdanie)
        self.button8.SetLabel(self.awyr)

    def zerowanie(self):
        #zerowanie danych w lekcji
        self.znam = []
        for i in range(self.maxslowwlekcji):
            self.znam.append(0)
        self.licznik = 1
        self.ilewiem = 0
        self.ilewiemprocent = 0
        self.napisy()
        if self.tryb == "ANG":
            self.cowidac = 'a'
        else:
            self.cowidac = 'pol'
        self.ktoryprzebieg = 1
        self.zaznaczindeks()
        self.mc.Hide()

    def playfilm(self, evt):
        #odtwarzanie filmu
        if self.comboBox:
            self.comboBox.Hide()
        if self.czyfilm == 'ok':
            if self.czyplay == 'nie':
                pygame.mixer.music.stop()
                self.mc.Show()
                self.mc.Play()
                self.czyplay = 'tak'
            else:
                self.mc.Stop()
                self.czyplay = 'nie'

    def OnEraseBackground(self, evt):
        # tlo
        dc = evt.GetDC()
        bmp = wx.Bitmap("tlon.bmp")
        if width > 1400:
            bmp = wx.Bitmap("tlon1.png")
        if width > 2000:
            bmp = wx.Bitmap("tlon2.png")
        dc.DrawBitmap(bmp, 0, 0)

    def animacjaradosc(self):
        # odtwarza animacje radość
        self.ctrl1.Play()
        self.timer.Start(3000, True)

    def animacjazlosc(self):
        # odtwarza animacje zlość
        self.ctrl.Play()
        self.timer.Start(3000, True)

    def animacjastop(self):
        #zerowanie animacji
        self.ctrl.Stop()
        self.ctrl1.Stop()
        self.ctrl1.Hide()
        self.ctrl.Hide()

    def OnTimer(self, evt):
        #odliczanie czasu animacji
        self.animacjastop()

    def wiemslowo(self):
        # inteligentne powtórki słowa
        if self.znam[self.licznik - 1] == 0:
            self.znam[self.licznik - 1] = 1
            self.ilewiem += 1

        if self.znam[self.licznik - 1] == 10:
            self.ilewiem += 0.5
            self.znam[self.licznik - 1] = 1

        if self.znam[self.licznik - 1] == 2:
            self.znam[self.licznik - 1] = 1
            self.ilewiem += 0.5

        if self.znam[self.licznik - 1] == 5:
            self.ilewiem += 0.5
            self.znam[self.licznik - 1] = 10

        if self.znam[self.licznik - 1] == 6:
            self.znam[self.licznik - 1] = 2
        self.przeliczwiedze()

    def przeliczwiedze(self):
        # oblicza ile wiesz w lekcji
        if self.ilewiem == 0:
            self.ilewiemprocent = 0
        self.ilewiemprocent = self.ilewiem * 100 / self.maxslowwlekcji
        self.ilewiemprocent = int(self.ilewiemprocent)
        self.napisy()

    def napisy(self):
        # aktualizacja danych
        self.nwiedza.SetLabel(" wiesz " + str(self.ilewiemprocent) + " %")
        self.nlicznik.SetLabel(" licznik " + str(self.licznik))
        self.nmaxslow.SetLabel("ile wyrazów w lekcji :" + str(self.maxslowwlekcji))
        # self.nznam.SetLabel(str(self.znam[0:] ))

    def szukajwolnego(self):
        # sprawdzanie pierwszego słowa w lekcji którego nie znasz
        if self.znam[self.licznik - 1] == 1:
            self.licznik += 1
            self.sprawdzlicznik()
            self.szukajwolnego()

    def sprawdzlicznik(self):
        #sprawdzenie czy wszyskie wyrazy w lekcji
        if self.licznik > self.maxslowwlekcji:
            self.zwiekszlek()
        self.nlicznik.SetLabel(" licznik " + str(self.licznik))


app = wx.App(0)
# obliczenie skali do ramek i napisów
width, height = wx.GetDisplaySize()
scale = (width / 1366)
# wywołanie programu
frame = wx.Frame(None, size=(width, height))
panel = TestPanel(frame)
frame.Show()
app.MainLoop()
