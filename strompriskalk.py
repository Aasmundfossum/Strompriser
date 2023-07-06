import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import datetime as datetime
import plotly.express as px
from plotly import graph_objects as go

st.set_page_config(
    page_title="Strømpriskalk",
    page_icon="🧮",
)

with open("styles/main.css") as f:
    st.markdown("<style>{}</style>".format(f.read()), unsafe_allow_html=True)

st.title('Strømpriskalkulator 🔌💸🧮')
st.markdown('Laget av Åsmund Fossum 👨🏼‍💻')
st.header('Inndata')

st.subheader('Forbruksdata')
FORBRUKSFIL = st.file_uploader(label='Timesdata for strømforbruk i enhet kW. Lenge 8760x1 ved normalår og 8784x1 ved skuddår.',type='xlsx')

st.subheader('Nettleiesatser')
konst_pris = st.checkbox(label='Bruk konstante verdier på nettleie og spotpris')

if konst_pris == True:
    c1, c2 = st.columns(2)
    with c1:
        konst_nettleie = st.number_input(label='Konstant timespris nettleie (kr/kWh)',value=0.50, step=0.01)
    with c2:
        konst_spot = st.number_input(label='Konstant timesverdi på spotpris (kr/kWh)',value=1.00, step=0.01)
    prissats_fil = 0
    type_kunde = 0
    sone = 0
    paaslag = 0
    spotprisfil_aar = 0
    mva = 0

elif konst_pris == False:
    prissats_fil = st.file_uploader(label='Fil på riktig format som inneholder prissatser for kapasitetsledd, energiledd og offentlige avgifter',type='xlsx')
    type_kunde = st.selectbox(label='Type strømkunde',options=['Privatkunde', 'Mindre næringskunde', 'Større næringskunde'],index=0)
    st.subheader('Spotpris')
    c1, c2, c3 = st.columns(3)
    with c1:
        sone = st.selectbox(label='Sone for spotpris',options=['NO1','NO2','NO3','NO4','NO5'],index=0)
    with c2:
        spotprisfil_aar = st.selectbox(label='Årstall for spotpriser',options=['2022', '2021', '2020'],index=0)
    with c3:
        paaslag = st.number_input(label='Påslag på spotpris (kr/kWh)', value=0.05, step=0.01)
    
    st.subheader('Merverdiavgift')
    c1, c2, c3 = st.columns(3)
    with c1:
        mva = st.checkbox(label='Priser inkludert mva.')
    
skuddaar = False
spotprisfil = 'Spotpriser.xlsx'

def bestem_prissatser(prissats_fil,type_kunde,mva):
    kap_sats = pd.read_excel(prissats_fil,sheet_name=type_kunde)
    
    if mva == False:
        energi = kap_sats.iloc[0,6]
        reduksjon_energi = kap_sats.iloc[1,6]
        fast_avgift = kap_sats.iloc[2,6]
    else:
        energi = kap_sats.iloc[0,7]
        reduksjon_energi = kap_sats.iloc[1,7]
        fast_avgift = kap_sats.iloc[2,7]

    starttid_reduksjon = kap_sats.iloc[0,10]                     # Klokkeslett for start av reduksjon i energileddpris
    sluttid_reduksjon = kap_sats.iloc[1,10]
    helg_spm = kap_sats.iloc[2,10]
    if helg_spm == 'Ja':
        helgereduksjon = True
    elif helg_spm == 'Nei':
        helgereduksjon = False
    
    max_kW_kap_sats = kap_sats.iloc[:,2]
    kap_sats = kap_sats.iloc[:,3]
    
    return max_kW_kap_sats,kap_sats,energi,reduksjon_energi,starttid_reduksjon,sluttid_reduksjon,fast_avgift,helgereduksjon

def fiks_forbruksfil(timesforbruksfil):
    forb = pd.read_excel(timesforbruksfil,sheet_name='Sheet1')
    forb = forb.to_numpy()
    forb = np.swapaxes(forb,0,1)
    forb = forb[0,:]
    return forb

def dager_i_hver_mnd(skuddaar):
    if skuddaar == True:
        dager_per_mnd = np.array([31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
    else:
        dager_per_mnd = np.array([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
    return dager_per_mnd

def ukedag_eller_helligdag(year, dag_nummer, helgereduksjon):               #dag_nummer er IKKE nullindeksert!
    if helgereduksjon == True:
        dato = datetime.datetime(year, 1, 1) + datetime.timedelta(dag_nummer - 1)
        if dato.weekday() >= 5:
            return "helg"
        if year == 2022:
            norske_helligdager = [
                (1, 1),  # New Year's Day
                (4, 14),
                (4, 15),
                (4, 17),
                (4, 18),
                (5, 1),  # Labor Day
                (5, 17),  # Constitution Day
                (5, 26),
                (6, 5),
                (6, 6),
                (12, 25),  # Christmas Day
                (12, 26),  # Boxing Day
        ]
        elif year == 2021:
            norske_helligdager = [
                (1, 1),  # New Year's Day
                (4, 1),
                (4, 2),
                (4, 4),
                (4, 5),
                (5, 1),  # Labor Day
                (5, 13),
                (5, 17),  # Constitution Day
                (5, 23),
                (5, 24),
                (12, 25),  # Christmas Day
                (12, 26),  # Boxing Day
        ]
        elif year == 2020:
            norske_helligdager = [
                (1, 1),  # New Year's Day
                (4, 9),
                (4, 10),
                (4, 12),
                (4, 13),
                (5, 1),  # Labor Day
                (5, 17),  # Constitution Day
                (5, 21),
                (5, 31),
                (6, 1),
                (12, 25),  # Christmas Day
                (12, 26),  # Boxing Day
        ]
        if (dato.month, dato.day) in norske_helligdager:
            return "helg"

        # If neither weekend nor holiday, consider it a regular weekday
        return "ukedag"
    elif helgereduksjon == False:                               # Hvis man ikke skal ha med helgereduksjon, regnes alle dager som ukedag.
        return 'ukedag'

def energiledd(forb,dager_per_mnd,spotprisfil_aar,helgereduksjon,energi,reduksjon_energi,starttid_reduksjon,sluttid_reduksjon):
    energiledd_time = np.zeros(len(forb))
    for i in range(24,len(forb)+24,24):
        dagsforb = forb[i-24:i]
        dagnr = i/24
        dagtype = ukedag_eller_helligdag(int(spotprisfil_aar),dagnr,helgereduksjon)

        if dagtype == 'ukedag':
            for j in range(0,len(dagsforb)):
                if j < sluttid_reduksjon or j >= starttid_reduksjon:
                    energiledd_time[i-24+j]=dagsforb[j]*(energi-reduksjon_energi)
                else:
                    energiledd_time[i-24+j]=dagsforb[j]*(energi)
        elif dagtype == 'helg':
            for j in range(0,len(dagsforb)):
                energiledd_time[i-24+j]=dagsforb[j]*(energi-reduksjon_energi)


    for j in range(0,len(forb)):                    #Setter energileddet lik 0 i timer med 0 forbruk
        if forb[j] == 0:
            energiledd_time[j] = 0

    energiledd_mnd = np.zeros(12)
    forrige = 0
    for k in range(0,len(dager_per_mnd)):
        energiledd_mnd[k] = np.sum(energiledd_time[forrige:forrige+dager_per_mnd[k]*24])
        forrige = forrige + dager_per_mnd[k]*24
    return energiledd_time,energiledd_mnd


def kapasitetsledd(forb,max_kW_kap_sats,kap_sats,dager_per_mnd):
    kapledd_mnd = np.zeros(12)
    kapledd_time = []
    forrige = 0
    forrige_mnd = 0
    for i in range(0,len(dager_per_mnd)):
        mnd_forb = forb[forrige:forrige+dager_per_mnd[i]*24]
        forrige = forrige + dager_per_mnd[i]*24

        forrige_dag = 0
        max_per_dag = np.zeros(dager_per_mnd[i])
        for j in range(0,dager_per_mnd[i]):
            dag_forb = mnd_forb[forrige_dag:forrige_dag+24]
            forrige_dag = forrige_dag + 24
            max_per_dag[j] = np.max(dag_forb)

        max_per_dag_sort = np.sort(max_per_dag)
        tre_hoyeste = max_per_dag_sort[-3:]
        snitt_tre_hoyeste = np.mean(tre_hoyeste)

        if snitt_tre_hoyeste == 0:                  # Hvis det ikke brukes noe strøm en måned, skal kapledd være 0 denne måneden
            kapledd_mnd[i] = 0
        else:
            for k in range(0,len(max_kW_kap_sats)):
                if snitt_tre_hoyeste < max_kW_kap_sats[k]:
                    break
            kapledd_mnd[i] = kap_sats[k]

        ant_timer_med_forb = 0
        for l in range(0,len(mnd_forb)):
            if mnd_forb[l] != 0:
                ant_timer_med_forb = ant_timer_med_forb+1

        for m in range(0,len(mnd_forb)):
            if mnd_forb[m] != 0:
                kapledd_time = kapledd_time + [kapledd_mnd[i]/(ant_timer_med_forb)]
            elif mnd_forb[m] == 0:
                kapledd_time = kapledd_time + [0]

        forrige_mnd = forrige_mnd + dager_per_mnd[i]*24
    kapledd_time = np.array(kapledd_time)
    
    return kapledd_time,kapledd_mnd

def offentlige_avgifter(forb,dager_per_mnd,fast_avgift):
    offentlig_mnd = np.zeros(12)
    forrige = 0
    for i in range(0,len(dager_per_mnd)):
        offentlig_mnd[i] = np.sum(forb[forrige:forrige+dager_per_mnd[i]*24])*fast_avgift
        forrige = forrige + dager_per_mnd[i]*24

    offentlig_time = forb*fast_avgift
    return offentlig_time,offentlig_mnd

def nettleie_hvis_konstant_sats(forb,dager_per_mnd,konst_nettleie):
    konst_nettleie_mnd = np.zeros(12)
    forrige = 0
    for i in range(0,len(dager_per_mnd)):
        konst_nettleie_mnd[i] = np.sum(forb[forrige:forrige+dager_per_mnd[i]*24])*konst_nettleie
        forrige = forrige + dager_per_mnd[i]*24

    konst_nettleie_time = forb*konst_nettleie
    return konst_nettleie_time,konst_nettleie_mnd

def spotpris(konst_pris,spotprisfil,spotprisfil_aar,sone,paaslag,forb,dager_per_mnd):
    if konst_pris == False:
        spot_sats = pd.read_excel(spotprisfil,sheet_name=spotprisfil_aar)
        spot_sats = spot_sats.loc[:,sone]+paaslag
        if mva == True:
            spot_time = forb*spot_sats
        elif mva == False:
            spot_time = forb*(spot_sats/1.25)
        spot_mnd = np.zeros(12)
        forrige = 0
        for k in range(0,len(dager_per_mnd)):
            spot_mnd[k] = np.sum(spot_time[forrige:forrige+dager_per_mnd[k]*24])
            forrige = forrige + dager_per_mnd[k]*24
        return spot_time,spot_mnd
    elif konst_pris == True:
        konst_spot_mnd = np.zeros(12)
        forrige = 0
        for i in range(0,len(dager_per_mnd)):
            konst_spot_mnd[i] = np.sum(forb[forrige:forrige+dager_per_mnd[i]*24])*konst_spot
            forrige = forrige + dager_per_mnd[i]*24

        konst_spot_time = forb*konst_spot
        return konst_spot_time,konst_spot_mnd

def nettleie_storre_naring(forb,dager_per_mnd):

    kap_sats = pd.read_excel(prissats_fil,sheet_name=type_kunde)
    fastledd_mnd = np.zeros(12)
    energiledd_mnd = np.zeros(12)
    fond_avgift_mnd = np.zeros(12)
    kapledd_mnd = np.zeros(12)
    offentlig_mnd = np.zeros(12)
    kapledd_time = []
    offentlig_time = []
    kol = 1

    fastledd_time = [kap_sats.iloc[0,kol]/(np.sum(dager_per_mnd)*24)] * (np.sum(dager_per_mnd)*24)
    fond_avgift_time = [kap_sats.iloc[10,kol]/(np.sum(dager_per_mnd)*24)] * (np.sum(dager_per_mnd)*24)
    energiledd_time = (kap_sats.iloc[2,kol]/100)*forb           # kr

    forrige = 0
    for i in range(0,len(dager_per_mnd)):
        mnd_forb = forb[forrige:forrige+dager_per_mnd[i]*24]
        energiledd_mnd[i] = np.sum(energiledd_time[forrige:forrige+dager_per_mnd[i]*24])
        forrige = forrige + dager_per_mnd[i]*24

        fastledd_mnd[i] = (kap_sats.iloc[0,kol]/(np.sum(dager_per_mnd)))*dager_per_mnd[i]
        fond_avgift_mnd[i] = (kap_sats.iloc[10,kol]/np.sum(dager_per_mnd))*dager_per_mnd[i]
        

        if 3 <= i <=8: 
            kapledd_time = kapledd_time + ([((kap_sats.iloc[5,kol]/(np.sum(dager_per_mnd)))*dager_per_mnd[i]*np.max(mnd_forb))/(dager_per_mnd[i]*24)] * dager_per_mnd[i]*24)
            kapledd_mnd[i] = np.sum([((kap_sats.iloc[5,kol]/(np.sum(dager_per_mnd)))*dager_per_mnd[i]*np.max(mnd_forb))/(dager_per_mnd[i]*24)] * dager_per_mnd[i]*24)
        else:
            kapledd_time = kapledd_time + ([((kap_sats.iloc[3,kol]/(np.sum(dager_per_mnd)))*dager_per_mnd[i]*np.max(mnd_forb))/(dager_per_mnd[i]*24)] * dager_per_mnd[i]*24)
            kapledd_mnd[i] = np.sum([((kap_sats.iloc[3,kol]/(np.sum(dager_per_mnd)))*dager_per_mnd[i]*np.max(mnd_forb))/(dager_per_mnd[i]*24)] * dager_per_mnd[i]*24)

        if i <=2:
            offentlig_time = offentlig_time + [kap_sats.iloc[7,kol]/100*mnd_forb]
            offentlig_mnd[i] = np.sum([kap_sats.iloc[7,kol]/100*mnd_forb])
        else:
            offentlig_time = offentlig_time + [kap_sats.iloc[8,kol]/100*mnd_forb]
            offentlig_mnd[i] = np.sum([kap_sats.iloc[8,kol]/100*mnd_forb])
    
    #kapledd_time = [item for sublist in kapledd_time for item in sublist]
    #kapledd_time = np.array(kapledd_time)
    offentlig_time = [item for sublist in offentlig_time for item in sublist]
    offentlig_time = np.array(offentlig_time)
    
    return energiledd_time,kapledd_time,offentlig_time,fastledd_time,fastledd_mnd,energiledd_time,energiledd_mnd,kapledd_time,kapledd_mnd,offentlig_time,offentlig_mnd,fond_avgift_time,fond_avgift_mnd


def hele_strompris(timesforbruksfil,konst_pris,prissats_fil,spotprisfil,spotprisfil_aar,sone,paaslag,type_kunde,mva,skuddaar):
    dager_per_mnd = dager_i_hver_mnd(skuddaar)
    forb = fiks_forbruksfil(timesforbruksfil)
    [spot_time,spot_mnd] = spotpris(konst_pris,spotprisfil,spotprisfil_aar,sone,paaslag,forb,dager_per_mnd)

    if konst_pris == False:
        if type_kunde == 'Større næringskunde':
            [energiledd_time,kapledd_time,offentlig_time,fastledd_time,fastledd_mnd,energiledd_time,energiledd_mnd,kapledd_time,kapledd_mnd,offentlig_time,offentlig_mnd,fond_avgift_time,fond_avgift_mnd
             ] = nettleie_storre_naring(forb,dager_per_mnd)
            tot_nettleie_time = fastledd_time+energiledd_time+kapledd_time+offentlig_time+fond_avgift_time
            tot_nettleie_mnd = fastledd_mnd+energiledd_mnd+kapledd_mnd+offentlig_mnd+fond_avgift_mnd
        
        else:
            [max_kW_kap_sats,kap_sats,energi,reduksjon_energi,starttid_reduksjon,sluttid_reduksjon,fast_avgift,helgereduksjon] = bestem_prissatser(prissats_fil,type_kunde,mva)
            [energiledd_time,energiledd_mnd] = energiledd(forb,dager_per_mnd,spotprisfil_aar,helgereduksjon,energi,reduksjon_energi,starttid_reduksjon,sluttid_reduksjon)
            [kapledd_time,kapledd_mnd] = kapasitetsledd(forb,max_kW_kap_sats,kap_sats,dager_per_mnd)
            [offentlig_time,offentlig_mnd] = offentlige_avgifter(forb,dager_per_mnd,fast_avgift)
            tot_nettleie_time = energiledd_time+kapledd_time+offentlig_time
            tot_nettleie_mnd = energiledd_mnd+kapledd_mnd+offentlig_mnd
            

    elif konst_pris == True:
        [tot_nettleie_time,tot_nettleie_mnd] = nettleie_hvis_konstant_sats(forb,dager_per_mnd,konst_nettleie)
        energiledd_time = 0
        kapledd_time = 0
        offentlig_time = 0
        plot_tittel = 'Strømpris med gitt forbruk, nettleie på '+str(konst_nettleie)+' kr/kWh og spotpris på '+str(konst_spot)+' kr/kWh'

    return energiledd_time,kapledd_time,offentlig_time,tot_nettleie_mnd,spot_mnd,tot_nettleie_time,spot_time,forb

   
def plot_resultater(energiledd_time,kapledd_time,offentlig_time,forb,tot_nettleie_mnd,spot_mnd,tot_nettleie_time,spot_time,type_kunde,sone):
    st.header('Resultater')
    if konst_pris == False:
        if mva == True:
            mva_str = 'inkl. mva.'
        elif mva == False:
            mva_str = 'ekskl. mva.'
        
        if type_kunde == 'Større næringskunde':
            timesnettleie_til_plot = pd.DataFrame({'Energiledd':energiledd_time, 'Effektledd':kapledd_time, 'Forbruksavgift':offentlig_time})
            fig1 = px.line(timesnettleie_til_plot, title='Nettleie for '+type_kunde.lower()+' med gitt forbruk '+mva_str, color_discrete_sequence=['#1d3c34', '#FFC358', '#48a23f'])
            fig1.update_layout(xaxis_title='Timer', yaxis_title='Timespris med gitt forbruk (kr)',legend_title=None)
            st.plotly_chart(fig1)
            
        else:
            timesnettleie_til_plot = pd.DataFrame({'Energiledd':energiledd_time, 'Kapasitetsledd':kapledd_time, 'Offentlige avgifter':offentlig_time})
            fig1 = px.line(timesnettleie_til_plot, title='Nettleie for '+type_kunde.lower()+' med gitt forbruk '+mva_str, color_discrete_sequence=['#1d3c34', '#FFC358', '#48a23f'])
            fig1.update_layout(xaxis_title='Timer', yaxis_title='Timespris med gitt forbruk (kr)',legend_title=None)
            st.plotly_chart(fig1)
        
        plot_tittel = 'Strømpris for '+type_kunde.lower()+' med gitt forbruk i '+sone+' basert på spotpriser i '+spotprisfil_aar+' '+mva_str
    elif konst_pris == True:
        plot_tittel = 'Strømpris med gitt forbruk og gitte satser på nettleie og spotpris'
    
    tot_nettleie_aar = np.sum(tot_nettleie_mnd)
    tot_strompris_aar = tot_nettleie_aar + np.sum(spot_mnd)
    tot_strompris_time = tot_nettleie_time+spot_time
    tot_forb = np.sum(forb)

    mnd_akse = ['Januar', 'Februar', 'Mars', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Desember']

    # Plotter nettleie og strømpris med timesoppløsning
    timestrompriser_til_plot = pd.DataFrame({"Total strømpris" : tot_strompris_time, "Spotpris" : spot_time, "Nettleie" : tot_nettleie_time})
    fig2 = px.line(timestrompriser_til_plot, title=plot_tittel, color_discrete_sequence=['#1d3c34', '#FFC358', '#48a23f'])
    fig2.update_layout(xaxis_title='Timer', yaxis_title='Timespris med gitt forbruk (kr)',legend_title=None)
    st.plotly_chart(fig2)


    # Plotter nettleie og strømpris med månedsoppløsning
    maanedsstrompriser_til_plot = pd.DataFrame({'Måned':mnd_akse,'Spotpris':spot_mnd,'Nettleie':tot_nettleie_mnd})
    fig5 = px.bar(maanedsstrompriser_til_plot,x='Måned',y=['Spotpris','Nettleie'],title=plot_tittel, color_discrete_sequence=['#FFC358', '#48a23f'])
    fig5.update_layout(yaxis_title='Månedspris med gitt forbruk (kr)',legend_title=None)
    st.plotly_chart(fig5)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric('Totalt forbruk:',f"{round(tot_forb)} kWh")
    with c2:
        st.metric('Total energikostnad dette året:',f"{round(tot_strompris_aar)} kr")
    with c3:
        st.metric('Gjennomsnittlig energikostnad per kWh',f"{round(tot_strompris_aar/tot_forb,3)} kr/kWh")


[energiledd_time,kapledd_time,offentlig_time,tot_nettleie_mnd,spot_mnd,tot_nettleie_time,spot_time,forb
 ] = hele_strompris(FORBRUKSFIL,konst_pris,prissats_fil,spotprisfil,spotprisfil_aar,sone,paaslag,type_kunde,mva,skuddaar)
plot_resultater(energiledd_time,kapledd_time,offentlig_time,forb,tot_nettleie_mnd,spot_mnd,tot_nettleie_time,spot_time,type_kunde,sone)

