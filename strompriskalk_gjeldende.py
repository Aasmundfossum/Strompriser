import numpy as np
import pandas as pd
import streamlit as st
import datetime as datetime
import plotly.express as px
from io import BytesIO
#from plotly import graph_objects as go

st.set_page_config(page_title="Strømpriskalk", page_icon="🧮")

with open("styles/main.css") as f:
    st.markdown("<style>{}</style>".format(f.read()), unsafe_allow_html=True)

####

class Strompriskalk:
    def __init__(self):
        pass

    def regn_ut_strompris(self):
        #Kjører hele beregningen, samt viser input og resultater i streamlit-nettside
        self.streamlit_input()
        if self.forbruksfil and (self.prissats_filnavn or self.konst_pris):
            self.bestem_prissatser()
            self.fiks_forbruksfil()
            self.dager_i_hver_mnd()
            self.spotpris()
            self.energiledd()
            self.kapasitetsledd()
            self.offentlige_avgifter()
            self.nettleie_hvis_konstant_sats()
            #self.spotpris()
            self.ekstra_nettleie_storre_naring()
            self.hele_nettleie()
            self.totaler()
            self.plot_resultater()
            self.last_ned_resultater()

    def streamlit_input(self):
        # Viser alle input-felt i streamlit
        st.title('Strømpriskalkulator 🔌💸🧮')
        st.markdown('Laget av Åsmund Fossum 👨🏼‍💻')
        st.header('Inndata')

        st.subheader('Forbruksdata')
        self.forbruksfil = st.file_uploader(label='Timesdata for strømforbruk i enhet kW med lengde 8760x1 og én tittel-rad.',type='xlsx')

        st.subheader('Nettleiesatser')
        self.konst_pris = st.checkbox(label='Bruk konstante verdier på nettleie og spotpris')
        if self.konst_pris == True:
            c1, c2 = st.columns(2)
            with c1:
                self.konst_nettleie = st.number_input(label='Konstant timespris nettleie (kr/kWh)',value=0.50, step=0.01)
            with c2:
                self.konst_spot = st.number_input(label='Konstant timesverdi på spotpris (kr/kWh)',value=1.00, step=0.01)
            
            self.prissats_filnavn = None

        elif self.konst_pris == False:
            #self.prissats_filnavn = st.file_uploader(label='Fil på riktig format som inneholder prissatser for kapasitetsledd, energiledd og offentlige avgifter',type='xlsx')
            self.nettleieselskap = st.selectbox('Velg nettleieselskap', ['BKK', 'Glitre', 'Lede', 'Tensio'])
            self.prissats_filnavn = f'Prissatser_nettleie_{self.nettleieselskap}.xlsx'
            
            self.type_kunde = st.selectbox(label='Type strømkunde',options=['Privatkunde', 'Mindre næringskunde', 'Større næringskunde'],index=0)
            st.subheader('Spotpris')
            c1, c2, c3 = st.columns(3)
            with c1:
                self.sone = st.selectbox(label='Sone for spotpris',options=['NO1','NO2','NO3','NO4','NO5'],index=0)
            with c2:
                self.spotprisfil_aar = st.selectbox(label='Årstall for spotpriser',options=['2023', '2022', '2021', '2020'],index=0)
            with c3:
                self.paaslag = st.number_input(label='Påslag på spotpris (kr/kWh)', value=0.05, step=0.01)
            
            st.subheader('Merverdiavgift')
            c1, c2, c3 = st.columns(3)
            with c1:
                self.mva = st.checkbox(label='Priser inkludert mva.')
                if self.mva == False:
                    mva_faktor = 1
                elif self.mva == True:
                    mva_faktor = 1.25

        self.plusskunde_modell = st.selectbox(label='Modell for håndtering av negativt forbruk (ikke helt ferdig)', options=['BKKs modell', 'Ingen'], index=0)
        self.skuddaar = False
        self.spotprisfil = 'Spotpriser.xlsx'
        self.mva_faktor = mva_faktor

    def bestem_prissatser(self):
        #Leser av prissatser for nettleie fra excel-filen som er lastet opp. Skjer kun hvis man ikke velger konstant pris
        if self.konst_pris == False:
            prissats_fil = pd.read_excel(self.prissats_filnavn, sheet_name=self.type_kunde)
            
            if self.type_kunde != 'Større næringskunde':
                
                if self.mva == False:
                    kol0 = 6
                    kol00 = 4
                else:
                    kol0 = 7
                    kol00 = 3
                self.energi = prissats_fil.iloc[0,kol0]                     #Energiledd inkl. fba.
                self.reduksjon_energi = prissats_fil.iloc[1,kol0]
                self.fast_avgift = prissats_fil.iloc[2,kol0]
                self.kap_sats = prissats_fil.iloc[:,kol00]

                self.max_kW_kap_sats = prissats_fil.iloc[:,2]
                self.starttid_reduksjon = prissats_fil.iloc[0,10]                     # Klokkeslett for start av reduksjon i energileddpris
                self.sluttid_reduksjon = prissats_fil.iloc[1,10]
                
                helg_spm = prissats_fil.iloc[2,10]
                if helg_spm == 'Ja':
                    self.helgereduksjon = True
                elif helg_spm == 'Nei':
                    self.helgereduksjon = False

            else:
                kol = 1
                self.energiledd_storre_naring_sommer = prissats_fil.iloc[2,kol]*self.mva_faktor
                self.energiledd_storre_naring_vinter = prissats_fil.iloc[3,kol]*self.mva_faktor
                self.kap_sats_apr_sept = prissats_fil.iloc[6,kol]*self.mva_faktor
                self.kap_sats_okt_mar = prissats_fil.iloc[4,kol]*self.mva_faktor
                self.avgift_sats_jan_mar = prissats_fil.iloc[8,kol]*self.mva_faktor
                self.avgift_sats_apr_des = prissats_fil.iloc[9,kol]*self.mva_faktor
                self.fastledd_sats = prissats_fil.iloc[0,kol]*self.mva_faktor
                self.sond_avgift_sats = prissats_fil.iloc[11,kol]*self.mva_faktor

            self.prissats_fil = prissats_fil


    def fiks_forbruksfil(self):
        # Gjør om timesforbruket i den opplastede filen til ønsket format
        #forb = pd.read_excel(self.forbruksfil,sheet_name='Sheet1')
        forb = pd.read_excel(self.forbruksfil)
        forb = forb.to_numpy()
        forb = np.swapaxes(forb,0,1)
        forb = forb[0,:]
        self.forb = forb

    def dager_i_hver_mnd(self):
        # Lager en vektor som inneholder antal dager i hver måned i året
        if self.skuddaar == True:
            self.dager_per_mnd = np.array([31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
        else:
            self.dager_per_mnd = np.array([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])

    def spotpris(self):
        # Hvis ikke konstante prisverdier: Leser av riktig kolonne (sone og år) i spotpristabell
        # Regner ut spotpris med timesoppløsning og månedsoppløsning. Hvis ikke konstante prisverdier tas hensyn til mva og påslag
        if self.konst_pris == False:
            spot_sats = pd.read_excel(self.spotprisfil,sheet_name=self.spotprisfil_aar)
            spot_sats = spot_sats.loc[:,self.sone]+self.paaslag
            spot_time = self.forb*(spot_sats/self.mva_faktor)
            spot_mnd = np.zeros(12)
            forrige = 0
            for k in range(0,len(self.dager_per_mnd)):
                spot_mnd[k] = np.sum(spot_time[forrige:forrige+self.dager_per_mnd[k]*24])
                forrige = forrige + self.dager_per_mnd[k]*24
            
            self.spot_time = spot_time
            self.spot_mnd = spot_mnd
        
        elif self.konst_pris == True:
            konst_spot_mnd = np.zeros(12)
            spot_sats = self.konst_spot
            forrige = 0
            for i in range(0,len(self.dager_per_mnd)):
                konst_spot_mnd[i] = np.sum(self.forb[forrige:forrige+self.dager_per_mnd[i]*24])*spot_sats
                forrige = forrige + self.dager_per_mnd[i]*24
            konst_spot_time = self.forb*spot_sats

            self.spot_time = konst_spot_time
            self.spot_mnd = konst_spot_mnd
        self.spot_sats = np.array(spot_sats)

    def energiledd(self):
        # Kun hvis ikke konstante prisverdier: 
        # Regner ut energiledd til nettleien basert på satsen som er lest av. 
        # Ved større næringskunde beregnes den på en annen måte enn ved enkelthusholdning og mindre næringskunde.
        # Returnerer energiledd med timesoppløsning og månedsoppløsning
        if self.konst_pris == False:
            energiledd_mnd = np.zeros(12)
            energiledd_time = []
            
            if self.type_kunde == 'Større næringskunde':
                forrige = 0
                for i in range(0,len(self.dager_per_mnd)):
                    if 3 <= i <=8: 
                        energiledd_sats = self.energiledd_storre_naring_sommer
                    else:
                        energiledd_sats = self.energiledd_storre_naring_vinter

                    forb_time_denne_mnd = self.forb[forrige:forrige+self.dager_per_mnd[i]*24]
                    energiledd_time_denne_mnd = (energiledd_sats/100)*forb_time_denne_mnd
                    spot_time_denne_mnd = self.spot_sats[forrige:forrige+self.dager_per_mnd[i]*24]
                    
                    for j in range(0,len(energiledd_time_denne_mnd)):
                        if energiledd_time_denne_mnd[j] < 0:
                            if self.plusskunde_modell == 'BKKs modell':
                                energiledd_time_denne_mnd[j] = 0.04*spot_time_denne_mnd[j]*forb_time_denne_mnd[j]           # "(...), betaler BKK 4 prosent av spotpris til plusskundene."
                            else:
                                energiledd_time_denne_mnd[j] = 0
                    
                    energiledd_time = energiledd_time + list((energiledd_time_denne_mnd))
                    
                    energiledd_mnd[i] = np.sum(energiledd_time_denne_mnd)
                    forrige = forrige + self.dager_per_mnd[i]*24

            else:
                energiledd_time = np.zeros(len(self.forb))
                for i in range(24,len(self.forb)+24,24):
                    dagsforb = self.forb[i-24:i]
                    dag_nummer = i/24
                    
                    #Ukedag eller helligdag:
                    dagtype ="ukedag"
                    if self.helgereduksjon == True:
                        year = int(self.spotprisfil_aar)
                        dato = datetime.datetime(year, 1, 1) + datetime.timedelta(dag_nummer - 1)
                        if dato.weekday() >= 5:         
                            dagtype="helg"
                        
                        if year == 2023:
                            norske_helligdager = [(1, 1), (4, 6), (4, 7), (4, 10), (5, 1), (5, 17), (5, 18), (5, 29), (12, 25), (12, 26),]
                        elif year == 2022:
                            norske_helligdager = [(1, 1), (4, 14), (4, 15), (4, 17), (4, 18), (5, 1), (5, 17), (5, 26), (6, 5), (6, 6), (12, 25), (12, 26),]
                        elif year == 2021:
                            norske_helligdager = [(1, 1), (4, 1), (4, 2), (4, 4), (4, 5), (5, 1), (5, 13), (5, 17), (5, 23), (5, 24), (12, 25), (12, 26),]
                        elif year == 2020:
                            norske_helligdager = [(1, 1), (4, 9), (4, 10), (4, 12), (4, 13), (5, 1), (5, 17), (5, 21), (5, 31), (6, 1), (12, 25), (12, 26),]
                        
                        if (dato.month, dato.day) in norske_helligdager:
                            dagtype="helg"
                    

                    if dagtype == 'ukedag':
                        for j in range(0,len(dagsforb)):
                            if j < self.sluttid_reduksjon or j >= self.starttid_reduksjon:
                                energiledd_time[i-24+j]=dagsforb[j]*(self.energi-self.reduksjon_energi)
                            else:
                                energiledd_time[i-24+j]=dagsforb[j]*(self.energi)
                    elif dagtype == 'helg':
                        for j in range(0,len(dagsforb)):
                            energiledd_time[i-24+j]=dagsforb[j]*(self.energi-self.reduksjon_energi)

                for j in range(0,len(self.forb)):                    #Setter energileddet lik 0 i timer med negativt
                    if self.forb[j] < 0:
                        energiledd_time[j] = 0

                forrige = 0
                for k in range(0,len(self.dager_per_mnd)):
                    energiledd_mnd[k] = np.sum(energiledd_time[forrige:forrige+self.dager_per_mnd[k]*24])
                    forrige = forrige + self.dager_per_mnd[k]*24
            
            self.energiledd_time = np.array(energiledd_time)
            self.energiledd_mnd = energiledd_mnd


    def kapasitetsledd(self):
        # Kun hvis ikke konstante prisverdier: 
        # Regner ut kapasitetsledd til nettleien basert på satsen som er lest av.
        # Hvis større næringskunde: effektledd = kapasitetsledd. Leser av sats fra fil og bruker denne 
        # Hvis mindre næringskunde eller enkelthusholdning: Finner passende kapasitetstrinn basert på tabell som ble avlest i bestem_prissatser. Bruker den for snitt tre høyeste døgnmakser
        # Returnerer kapasitetsledd med timesoppløsning og månedsoppløsning
        if self.konst_pris == False:
            kapledd_mnd = np.zeros(12)
            kapledd_time = []
            forrige = 0
            for i in range(0,len(self.dager_per_mnd)):
                mnd_forb = self.forb[forrige:forrige+self.dager_per_mnd[i]*24]
                forrige = forrige + self.dager_per_mnd[i]*24

                if self.type_kunde == 'Større næringskunde':   
                    if 3 <= i <=8: 
                        kap_sats = self.kap_sats_apr_sept
                    else:
                        kap_sats = self.kap_sats_okt_mar
                    
                    kapledd_time_denne_mnd = [((kap_sats/(np.sum(self.dager_per_mnd)))*self.dager_per_mnd[i]*np.max(mnd_forb))/(self.dager_per_mnd[i]*24)] * self.dager_per_mnd[i]*24
                    for j in range(0,len(kapledd_time_denne_mnd)):
                        if kapledd_time_denne_mnd[j] < 0:
                            kapledd_time_denne_mnd[j] = 0
                    kapledd_time = kapledd_time + (kapledd_time_denne_mnd)
                    kapledd_mnd[i] = np.sum(kapledd_time_denne_mnd)

                else:
                    forrige_dag = 0
                    forrige_mnd = 0
                    max_per_dag = np.zeros(self.dager_per_mnd[i])
                    for j in range(0,self.dager_per_mnd[i]):
                        dag_forb = mnd_forb[forrige_dag:forrige_dag+24]
                        forrige_dag = forrige_dag + 24
                        max_per_dag[j] = np.max(dag_forb)

                    max_per_dag_sort = np.sort(max_per_dag)
                    tre_hoyeste = max_per_dag_sort[-3:]
                    snitt_tre_hoyeste = np.mean(tre_hoyeste)

                    for k in range(0,len(self.max_kW_kap_sats)):
                        if snitt_tre_hoyeste < self.max_kW_kap_sats[k]:
                            break
                    kapledd_mnd[i] = self.kap_sats[k]

                    for m in range(0,len(mnd_forb)):
                        kapledd_time = kapledd_time + [kapledd_mnd[i]/(len(mnd_forb))]          # Fordeler månedens kapasitetsledd (kr/mnd) utover alle månedens timer

                    forrige_mnd = forrige_mnd + self.dager_per_mnd[i]*24
            
            kapledd_time = np.array(kapledd_time)

            self.kapledd_time = np.array(kapledd_time)
            self.kapledd_mnd = kapledd_mnd

    def offentlige_avgifter(self):
        # Kun hvis ikke konstante prisverdier: 
        # Regner ut ekstra offentlige avgifter i nettleien basert på satsen som er lest av.
        # Hvis større næringskunde: Forbruksavgift (da denne ikke er inkludert noe annet sted) kr/kWh
        # Hvis mindre næringskunde: Avlest sats på fast avgift (satt til 0)
        # Hvis privatkunde: Avlest sats på fast avgift (enovaavgift) kr/kWh
        # Returnerer offentlige avgifter med timesoppløsning og månedsoppløsning
        if self.konst_pris == False:
            if self.type_kunde == 'Større næringskunde':
                offentlig_mnd = np.zeros(12)
                offentlig_time = []
                forrige = 0
                for i in range(0,len(self.dager_per_mnd)):
                    mnd_forb = self.forb[forrige:forrige+self.dager_per_mnd[i]*24]
                    forrige = forrige + self.dager_per_mnd[i]*24
                    if i <=2:
                        avgift_sats = self.avgift_sats_jan_mar
                    else:
                        avgift_sats = self.avgift_sats_apr_des
                    
                    offentlig_time = offentlig_time + [avgift_sats/100*mnd_forb]
                    offentlig_mnd[i] = np.sum([avgift_sats/100*mnd_forb])

                offentlig_time = [item for sublist in offentlig_time for item in sublist]
                offentlig_time = np.array(offentlig_time)
                for j in range(0,len(offentlig_time)):
                    if offentlig_time[j] < 0:
                        offentlig_time[j] = 0
                for k in range(0,len(offentlig_mnd)):
                    if offentlig_mnd[k] < 0:
                        offentlig_mnd[k] = 0
            else:            
                offentlig_time = self.forb*self.fast_avgift
            
                for i in range(0,len(offentlig_time)):
                    if offentlig_time[i] < 0:
                        offentlig_time[i] = 0

                offentlig_mnd = np.zeros(12)    
                forrige = 0
                for k in range(0,len(self.dager_per_mnd)):
                    offentlig_mnd[k] = np.sum(offentlig_time[forrige:forrige+self.dager_per_mnd[k]*24])
                    forrige = forrige + self.dager_per_mnd[k]*24

            self.offentlig_time = offentlig_time
            self.offentlig_mnd = offentlig_mnd

    def nettleie_hvis_konstant_sats(self):
        # Kun hvis konstante prisverdier
        # Regner ut total nettleie med timesoppløsning og månedsoppløsning
        if self.konst_pris == True:
            konst_nettleie_mnd = np.zeros(12)
            forrige = 0
            for i in range(0,len(self.dager_per_mnd)):
                konst_nettleie_mnd[i] = np.sum(self.forb[forrige:forrige+self.dager_per_mnd[i]*24])*self.konst_nettleie
                forrige = forrige + self.dager_per_mnd[i]*24

            konst_nettleie_time = self.forb*self.konst_nettleie
            
            self.konst_nettleie_time = konst_nettleie_time
            self.konst_nettleie_mnd = konst_nettleie_mnd

    def ekstra_nettleie_storre_naring(self):
        # Kun hvis ikke konstante prisverdier:
        # Leser av og Regner ut ekstra deler av nettleien som kun finnes for større næringskunder: Fastledd og næringsavgift til energifondet
        # Fordeler denne kun mellom alle årets timer
        if self.konst_pris == False:
            if self.type_kunde == 'Større næringskunde':
                fastledd_mnd = np.zeros(12)
                fond_avgift_mnd = np.zeros(12)
                fastledd_time = []
                fond_avgift_time =[]

                forrige = 0
                for i in range(0,len(self.dager_per_mnd)):
                    mnd_forb = self.forb[forrige:forrige+self.dager_per_mnd[i]*24]
                    forrige = forrige + self.dager_per_mnd[i]*24  

                    fastledd_mnd[i] = (self.fastledd_sats/(np.sum(self.dager_per_mnd)))*self.dager_per_mnd[i]
                    fond_avgift_mnd[i] = (self.sond_avgift_sats/np.sum(self.dager_per_mnd))*self.dager_per_mnd[i]

                    for m in range(0,len(mnd_forb)):
                        fastledd_time = fastledd_time + [fastledd_mnd[i]/(len(mnd_forb))]
                        fond_avgift_time = fond_avgift_time + [fond_avgift_mnd[i]/(len(mnd_forb))]

                self.fastledd_time = fastledd_time
                self.fastledd_mnd = fastledd_mnd
                self.fond_avgift_time = fond_avgift_time
                self.fond_avgift_mnd = fond_avgift_mnd 

    def hele_nettleie(self):
        # Regner ut total nettleie som sum av de ulike delene, avhengige av hvilke som er aktuelle i ulike tilfeller
        if self.konst_pris == False:
            if self.type_kunde == 'Større næringskunde':
                tot_nettleie_time = self.fastledd_time+self.energiledd_time+self.kapledd_time+self.offentlig_time+self.fond_avgift_time
                tot_nettleie_mnd = self.fastledd_mnd+self.energiledd_mnd+self.kapledd_mnd+self.offentlig_mnd+self.fond_avgift_mnd                 
            
            else:
                tot_nettleie_time = self.energiledd_time+self.kapledd_time+self.offentlig_time
                tot_nettleie_mnd = self.energiledd_mnd+self.kapledd_mnd+self.offentlig_mnd
                
        elif self.konst_pris == True:
            tot_nettleie_time = self.konst_nettleie_time
            tot_nettleie_mnd = self.konst_nettleie_mnd
            self.plot_tittel = 'Strømpris , nettleie på '+str(self.konst_nettleie)+' kr/kWh og spotpris på '+str(self.konst_spot)+' kr/kWh'

        self.tot_nettleie_time = tot_nettleie_time
        self.tot_nettleie_mnd = tot_nettleie_mnd

    def totaler(self):
        # Regner ut total strømpris som sum av total nettleie og spotpris
        tot_nettleie_aar = np.sum(self.tot_nettleie_mnd)
        self.tot_strompris_aar = tot_nettleie_aar + np.sum(self.spot_mnd)
        self.tot_strompris_time = self.tot_nettleie_time+self.spot_time
        self.tot_forb = np.sum(self.forb)

    def plot_resultater(self):
        # Skriver ut og plotter alle resultater til streamlit-nettsiden
        # Hvilke ting som plottes i figurene er avhengig av valg i input.
        st.header('Resultater')

        if self.konst_pris == False:
            if self.mva == True:
                mva_str = 'inkl. mva.'
            elif self.mva == False:
                mva_str = 'ekskl. mva.'
            
            if self.type_kunde == 'Større næringskunde':
                timesnettleie_til_plot = pd.DataFrame({'Energiledd':self.energiledd_time, 'Effektledd':self.kapledd_time, 'Forbruksavgift':self.offentlig_time, 'Fastledd':self.fastledd_time, 'Avgift til energifondet':self.fond_avgift_time})
                plot_farger = ['#1d3c34', '#FFC358', '#48a23f', '#b7dc8f', '#FAE3B4']
            else:
                timesnettleie_til_plot = pd.DataFrame({'Energiledd inkl. fba.':self.energiledd_time, 'Kapasitetsledd':self.kapledd_time, 'Andre offentlige avgifter':self.offentlig_time})
                plot_farger = ['#1d3c34', '#FFC358', '#48a23f']

            fig1 = px.line(timesnettleie_til_plot, title='Nettleie for '+self.type_kunde.lower()+' '+mva_str, color_discrete_sequence=plot_farger)
            fig1.update_layout(xaxis_title='Timer', yaxis_title='Timespris  (kr)',legend_title=None)
            st.plotly_chart(fig1)  
            
            plot_tittel = 'Strømpris for '+self.type_kunde.lower()+'  i '+self.sone+' basert på spotpriser i '+self.spotprisfil_aar+' '+mva_str
        
        elif self.konst_pris == True:
            plot_tittel = 'Strømpris  og gitte satser på nettleie og spotpris'
        
        # Plotter nettleie og strømpris med timesoppløsning
        timestrompriser_til_plot = pd.DataFrame({"Total strømpris" : self.tot_strompris_time, "Spotpris" : self.spot_time, "Nettleie" : self.tot_nettleie_time})
        fig2 = px.line(timestrompriser_til_plot, title=plot_tittel, color_discrete_sequence=['#1d3c34', '#FFC358', '#48a23f'])
        fig2.update_layout(xaxis_title='Timer', yaxis_title='Timespris  (kr)',legend_title=None)
        st.plotly_chart(fig2)


        # Plotter nettleie og strømpris med månedsoppløsning
        mnd_akse = ['Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Desember']
        maanedsstrompriser_til_plot = pd.DataFrame({'Måned':mnd_akse,'Spotpris':self.spot_mnd,'Nettleie':self.tot_nettleie_mnd})
        fig5 = px.bar(maanedsstrompriser_til_plot,x='Måned',y=['Spotpris','Nettleie'],title=plot_tittel, color_discrete_sequence=['#FFC358', '#48a23f'])
        fig5.update_layout(yaxis_title='Månedspris  (kr)',legend_title=None)
        st.plotly_chart(fig5)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric('Totalt forbruk:',f'{"{:,}".format(round(self.tot_forb)).replace(",", " ")} kWh')
        with c2:
            st.metric(f'Total energikostnad i {self.spotprisfil_aar}:',f'{"{:,}".format(round(self.tot_strompris_aar)).replace(",", " ")} kr')
        with c3:
            st.metric('Gjennomsnittlig energikostnad per kWh',f"{round(self.tot_strompris_aar/self.tot_forb,3)} kr/kWh")

        self.timestrompriser_til_plot = timestrompriser_til_plot
        self.maanedsstrompriser_til_plot = maanedsstrompriser_til_plot

    def last_ned_resultater(self):
        @st.cache_data
        def convert_df_to_excel(df):
            # Convert the DataFrame to an Excel file in-memory
            excel_file = BytesIO()
            with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=True, sheet_name="Sheet1")
                #writer.save()
            excel_file.seek(0)  # Set cursor to the beginning of the file
            return excel_file

        @st.fragment
        def download_button_hour():
            st.download_button(
                label="Last ned timesdata som Excel-fil",
                data=timesopplost_til_excel,
                file_name="Strømpriser_timesoppløst.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        @st.fragment
        def download_button_month():
            st.download_button(
                label="Last ned månedsdata som Excel-fil",
                data=maanedsopplost_til_excel,
                file_name="Strømpriser_månedsoppløst.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # Replace final_dataframe_file with the Excel version
        timesopplost_til_excel = convert_df_to_excel(self.timestrompriser_til_plot)
        maanedsopplost_til_excel = convert_df_to_excel(self.maanedsstrompriser_til_plot)

        st.markdown('---')
        c1, c2 = st.columns(2)
        with c1:
            download_button_hour()
        with c2:
            download_button_month()
        

Strompriskalk().regn_ut_strompris()

# Mulige forbedringer/tillegg:
# Forbruksavgift er nå inkludert i energiledd for privatkunder og mindre næringskunder, men den bør kanskje skilles ut og plasseres i kategorien "offentlige avgifter" sammen med Enovaavgift.
# Mindre næringskunder skal muligens også betale Enovapåslag på 800 kr/år, men dette er kun lagt inn for større næringskunder.
# Excel-filene med prissatser kan eventuelt bygges inn i koden i stedet for at de skal lastes opp. Kan også legge til prissatser fra flere nettselskaper.
# Bør kanskje legges inn noe som gjør at koden fungerer selv om timesforbruket er fra et skuddår.
