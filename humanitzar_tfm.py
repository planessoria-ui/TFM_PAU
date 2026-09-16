# -*- coding: utf-8 -*-
"""
Humanitza el text i corregeix l'estructura del TFM (v2 -> v3), editant el
document ORIGINAL de Pau per preservar-ne totes les imatges i el format.

Canvis:
  - Elimina patrons d'escriptura d'IA (guions llargs, "clau/valuós/robust",
    significació inflada, tancaments d'una línia, staging).
  - Corregeix frases trencades als resums ES i EN.
  - Arregla la numeració de figures i taules (camps SEQ amb resultat en memòria).
  - Afegeix la secció de sensors i renumera 4.3-4.7 -> 4.4-4.8.
  - Resol incoherències (nombre de vols, terminologia poda/maneig de capçada).
"""
import copy
import docx
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/2145bd8a-TFM_Pau_Planes_IRTA_v2.docx"
DST = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v3.docx"

doc = docx.Document(SRC)
paras = doc.paragraphs


def find(prefix, start=0):
    """Índex del primer paràgraf que comenci pel prefix donat."""
    for i in range(start, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip().startswith(prefix):
            return i
    raise LookupError(prefix[:70])


def set_text(p, new):
    """Reemplaça el text conservant el format del primer run."""
    runs = p.runs
    if not runs:
        p.add_run(new)
        return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def replace(prefix, new):
    i = find(prefix)
    set_text(doc.paragraphs[i], new)
    return i


# ---------------------------------------------------------------------------
# 1. RESUM / RESUMEN / ABSTRACT
# ---------------------------------------------------------------------------
replace("La gestió eficient de l'aigua de reg",
        "La gestió eficient de l'aigua de reg és un dels principals reptes de la "
        "viticultura mediterrània en un context d'escassetat hídrica i de canvi "
        "climàtic. Per ajustar el reg a la demanda real del cultiu cal conèixer amb "
        "precisió l'evapotranspiració real (ETa) a escala de parcel·la. El maneig de "
        "capçada modifica el volum del dosser, l'índex d'àrea foliar (LAI) i la "
        "fracció de radiació fotosintèticament activa interceptada (fIPAR), i amb "
        "això condiciona la transpiració i el balanç energètic de la vinya.")

replace("Aquest Treball Final de Màster, desenvolupat",
        "Aquest Treball Final de Màster, desenvolupat en col·laboració amb l'IRTA, "
        "avalua l'efecte dels tractaments de maneig de capçada sobre "
        "l'evapotranspiració de la vinya combinant teledetecció d'alta precisió amb "
        "dron i mesures de radiació PAR en camp. L'assaig es fa en una parcel·la de "
        "la varietat 'Ull de Llebre' (Tempranillo) a Raimat, amb un disseny en blocs "
        "de cinc tractaments d'esporga i quatre repeticions. Es faran quatre vols de "
        "dron al llarg del cicle vegetatiu amb càmeres tèrmica i multiespectral. De "
        "les imatges se'n deriven les variables biofísiques del dosser (LAI, fIPAR, "
        "alçada i volum) i les temperatures de dosser i de sòl, que alimenten els "
        "models de balanç energètic de dues fonts (TSEB) i de Shuttleworth–Wallace "
        "per estimar l'ETa i l'ET potencial. L'índex d'estrès hídric del cultiu "
        "(CWSI = 1 − ETa/ETp) es relaciona amb mesures de potencial hídric de tija "
        "(Ψstem), i les estimacions de fIPAR es validen amb mesures de ceptòmetre.")

replace("Paraules clau: vinya; evapotranspiració",
        "Paraules clau: vinya; Tempranillo; evapotranspiració; teledetecció; dron; "
        "radiació PAR; maneig de capçada; estrès hídric; reg de precisió.")

replace("La gestión eficiente del agua de riego",
        "La gestión eficiente del agua de riego es uno de los principales retos de la "
        "viticultura mediterránea en un contexto de escasez hídrica y de cambio "
        "climático. Para ajustar el riego a la demanda real del cultivo hay que "
        "conocer con precisión la evapotranspiración real (ETa) a escala de parcela. "
        "El manejo de copa modifica el volumen del dosel, el índice de área foliar "
        "(LAI) y la fracción de radiación fotosintéticamente activa interceptada "
        "(fIPAR), y con ello condiciona la transpiración y el balance energético de "
        "la viña.")

replace("Este Trabajo Final de Máster, desarrollado",
        "Este Trabajo Final de Máster, desarrollado en colaboración con el IRTA, "
        "evalúa el efecto de los tratamientos de manejo de copa sobre la "
        "evapotranspiración de la viña combinando teledetección de alta precisión con "
        "dron y medidas de radiación PAR en campo. El ensayo se realiza en una "
        "parcela de la variedad 'Tempranillo' en Raimat, con un diseño en bloques de "
        "cinco tratamientos de poda y cuatro repeticiones. Se harán cuatro vuelos de "
        "dron a lo largo del ciclo vegetativo con cámaras térmica y multiespectral. De "
        "las imágenes se derivan las variables biofísicas del dosel (LAI, fIPAR, "
        "altura y volumen) y las temperaturas de dosel y de suelo, que alimentan los "
        "modelos de balance energético de dos fuentes (TSEB) y de Shuttleworth–Wallace "
        "para estimar la ETa y la ET potencial. El índice de estrés hídrico del "
        "cultivo (CWSI = 1 − ETa/ETp) se relaciona con medidas de potencial hídrico de "
        "tallo (Ψstem), y las estimaciones de fIPAR se validan con medidas de "
        "ceptómetro.")

replace("Palabras clave: viña; evapotranspiración",
        "Palabras clave: viña; Tempranillo; evapotranspiración; teledetección; dron; "
        "radiación PAR; manejo de copa; estrés hídrico; riego de precisión.")

replace("Efficient irrigation water management",
        "Efficient irrigation water management is one of the main challenges of "
        "Mediterranean viticulture under water scarcity and climate change. Matching "
        "irrigation to the real crop demand requires an accurate estimate of actual "
        "evapotranspiration (ETa) at the plot scale. Canopy management changes canopy "
        "volume, leaf area index (LAI) and the fraction of intercepted "
        "photosynthetically active radiation (fIPAR), and through them it conditions "
        "vineyard transpiration and energy balance.")

replace("This Master's Thesis, carried out",
        "This Master's Thesis, carried out in collaboration with IRTA, assesses the "
        "effect of canopy-management treatments on vineyard evapotranspiration by "
        "combining high-resolution UAV remote sensing and field PAR measurements. The "
        "trial takes place in a 'Tempranillo' plot in Raimat, with a block design of "
        "five pruning treatments and four replicates. Four UAV flights will be flown "
        "over the growing season with thermal and multispectral cameras. Canopy "
        "biophysical variables (LAI, fIPAR, height and volume) and canopy and soil "
        "temperatures are derived from the imagery and used to drive the two-source "
        "energy balance (TSEB) and Shuttleworth–Wallace models to estimate ETa and "
        "potential ET. The crop water stress index (CWSI = 1 − ETa/ETp) is related to "
        "stem water potential (Ψstem) measurements, and fIPAR estimates are validated "
        "against ceptometer readings.")

replace("Key words: vineyard; evapotranspiration",
        "Key words: vineyard; Tempranillo; evapotranspiration; remote sensing; UAV; "
        "PAR radiation; canopy management; water stress; precision irrigation.")

# ---------------------------------------------------------------------------
# 2. INTRODUCCIÓ
# ---------------------------------------------------------------------------
replace("L'aigua és el principal factor limitant",
        "L'aigua és el principal factor limitant de la producció agrària a la conca "
        "mediterrània. Els escenaris de canvi climàtic projecten per al sud d'Europa "
        "un augment de la temperatura, més episodis de sequera i menys aigua "
        "disponible per al reg. La viticultura mediterrània, un sector de pes econòmic "
        "i paisatgístic a Catalunya, ha d'anar per tant cap a estratègies de reg que "
        "ajustin els aports hídrics a la demanda de cada parcel·la.")

replace("La poda és una de les pràctiques de maneig",
        "El maneig de capçada és una de les pràctiques que més condiciona "
        "l'arquitectura del dosser de la vinya. En determinar el nombre de gemmes i "
        "de pàmpols, la poda regula el desenvolupament de l'àrea foliar, la mida i el "
        "volum de la capçada i, amb això, la radiació interceptada i la superfície "
        "transpirant (Choné et al., 2001). Intensitats de poda diferents donen dossers "
        "amb capacitats diferents d'interceptar radiació i d'evapotranspirar, però "
        "aquest efecte encara s'ha quantificat poc en vinya amb mètodes operatius i "
        "d'alta resolució.")

replace("Concretament, el treball s'emmarca",
        "El treball s'emmarca en l'assaig de maneig de capçada que el Programa d'Ús "
        "Eficient de l'Aigua de l'IRTA desenvolupa en una parcel·la de la varietat "
        "'Ull de Llebre' (Tempranillo) a Raimat, dins dels projectes VITIMPACT i "
        "ADAPTEX, per augmentar la resiliència de la vinya davant del canvi climàtic. "
        "En aquest assaig, cinc nivells d'esporga generen capçades de mides "
        "contrastades. En campanyes anteriors ja s'hi van observar diferències de "
        "consum d'aigua de reg entre tractaments: les capçades més grans van arribar a "
        "consumir un 14 % més d'aigua que el maneig comercial, mentre que l'esporga "
        "severa el va reduir prop d'un 20 %. Aquest TFM aborda una de les línies de "
        "treball previstes en l'assaig, la quantificació de la transpiració i de "
        "l'evapotranspiració per teledetecció, que complementa les mesures directes "
        "amb sensors de flux de saba (sap flow sensors) i el seguiment del potencial "
        "hídric de tija.")

replace("El treball s'organitza en vuit capítols",
        "El treball s'organitza en vuit capítols. Després d'aquesta introducció, el "
        "capítol 2 desenvolupa el marc teòric: els fonaments de l'evapotranspiració, "
        "els models de balanç energètic, les variables biofísiques del dosser, la "
        "teledetecció amb dron, els mètodes de mesura de la radiació PAR, l'efecte de "
        "la poda i els indicadors d'estrès hídric. El capítol 3 presenta els objectius "
        "i les hipòtesis. El capítol 4 detalla els materials i mètodes: àrea d'estudi, "
        "disseny experimental, sensors instal·lats, vols de dron, processament "
        "d'imatges, mesures de camp, models d'ET i anàlisi estadística. Els capítols 5 "
        "i 6 recullen els resultats i la seva discussió. El capítol 7 sintetitza les "
        "conclusions i el capítol 8 llista les referències bibliogràfiques.")

# ---------------------------------------------------------------------------
# 3. MARC TEÒRIC
# ---------------------------------------------------------------------------
replace("El reg deficitari es defineix",
        "El reg deficitari consisteix a aplicar menys aigua que els requeriments "
        "complets d'evapotranspiració del cultiu, per reduir el consum amb una "
        "penalització mínima sobre la qualitat de la collita, o fins i tot amb un "
        "benefici (Fereres i Soriano, 2007). En vinya funciona bé perquè un dèficit "
        "hídric controlat afavoreix l'acumulació de compostos fenòlics i la "
        "concentració del most. Ara bé, aplicar-lo amb precisió exigeix conèixer, "
        "parcel·la a parcel·la, l'evapotranspiració real del cultiu i el grau d'estrès "
        "de la planta, i d'aquí ve l'aproximació metodològica d'aquest treball.")

replace("El model de balanç energètic de dues fonts (Two-Source",
        "El model de balanç energètic de dues fonts (Two-Source Energy Balance, TSEB), "
        "formulat per Norman et al. (1995) i desenvolupat després per Kustas i Norman, "
        "tracta el sòl i la vegetació com dues fonts diferenciades d'energia i de "
        "temperatura. A partir de la temperatura radiomètrica de la superfície, "
        "l'índex d'àrea foliar i les dades meteorològiques, el TSEB separa la "
        "transpiració del dosser de l'evaporació del sòl. Això importa especialment en "
        "cultius llenyosos en files com la vinya, on el sòl descobert ocupa una "
        "fracció important de la superfície, i és la raó per la qual el TSEB "
        "s'utilitza habitualment per estimar l'ETa en fruiters i vinya amb imatges "
        "d'alta resolució de dron (Bellvert et al., 2021). La implementació està "
        "disponible en obert al paquet pyTSEB (Nieto i Kustas). Les avaluacions del "
        "model mostren que les estimacions combinades dels fluxos de calor del sòl i "
        "de la vegetació s'ajusten a les observacions amb errors de l'ordre del 20 %, "
        "cosa que en sosté l'aplicació en cobertes parcials com els cultius en files "
        "(Kustas i Norman, 1999).")

replace("El funcionament del dosser i el seu consum",
        "El funcionament del dosser i el seu consum d'aigua depenen d'unes quantes "
        "variables biofísiques. L'índex d'àrea foliar (LAI) és la superfície foliar "
        "per unitat de superfície de sòl i determina la capacitat fotosintètica i "
        "transpirant del cultiu. La fracció de radiació fotosintèticament activa "
        "interceptada (fIPAR) quantifica la proporció de radiació PAR que el dosser "
        "intercepta, i depèn del LAI i de l'estructura de la coberta. Totes dues són "
        "entrades dels models de balanç energètic i, alhora, descriuen directament "
        "l'efecte de la poda sobre el dosser (Bellvert et al., 2021; Belaid et al., "
        "2025).")

replace("En cultius en files com la vinya, la fracció",
        "En cultius en files com la vinya, la fracció de coberta vegetal i el volum de "
        "la capçada condicionen tant la radiació interceptada com la partició entre "
        "transpiració i evaporació del sòl. Per això cal caracteritzar-los bé, amb "
        "mètodes òptics de camp o amb teledetecció, abans d'estimar l'ET.")

replace("Els vehicles aeris no tripulats",
        "Els vehicles aeris no tripulats (UAV o drons) han canviat la manera de "
        "monitorar els cultius, perquè permeten adquirir imatges amb resolucions "
        "espacials centimètriques i amb molta flexibilitat temporal. Amb càmeres "
        "multiespectrals, que registren bandes del visible, el red-edge i l'infraroig "
        "proper, es poden calcular índexs de vegetació com el NDVI i estimar el LAI i "
        "la fIPAR. Les càmeres tèrmiques mesuren la temperatura de la superfície, "
        "necessària per als models de balanç energètic i per als indicadors d'estrès "
        "hídric (Bellvert et al., 2014). El processament fotogramètric de les imatges "
        "(structure-from-motion) genera, a més, models digitals de superfície i del "
        "terreny, a partir dels quals es deriven l'alçada i el volum de la capçada de "
        "cada cep.")

replace("L'alta resolució de les imatges de dron",
        "L'alta resolució de les imatges de dron és especialment útil en vinya, on "
        "permet distingir els píxels purs de vegetació dels de sòl i obtenir així "
        "temperatures de dosser (Tc) i de sòl (Ts) diferenciades. Això millora la "
        "partició de fluxos del model TSEB respecte dels sensors satel·litaris de "
        "menor resolució (Bellvert et al., 2021).")

replace("La mesura de referència de la fIPAR en camp",
        "La mesura de referència de la fIPAR en camp es fa amb ceptòmetres lineals, "
        "com l'AccuPAR LP-80, que comparen la radiació PAR per sobre i per sota del "
        "dosser. La mesura puntual al migdia solar, però, no captura la variació "
        "diürna de la interceptació de llum, que en cultius en files depèn molt de "
        "l'orientació de les fileres, de la geometria del dosser i de l'angle solar. "
        "Per superar aquesta limitació, Belaid et al. (2025) han proposat i validat un "
        "mètode basat en imatges hemisfèriques preses amb una càmera d'acció de baix "
        "cost (GoPro), que reconstrueix la corba diürna de fIPAR i el seu valor "
        "integrat diari amb un grau d'acord elevat respecte del ceptòmetre. És un "
        "mètode operatiu i econòmic, aplicable en vinya i en fruiters. La Figura 2 "
        "mostra per què una única mesura al migdia pot subestimar la interceptació "
        "diària de llum en un cultiu en files.")

replace("La poda regula el balanç entre vigor",
        "La poda regula el balanç entre vigor vegetatiu i càrrega productiva del cep. "
        "En determinar el nombre de gemmes que brotaran, condiciona el desenvolupament "
        "de l'àrea foliar, la mida de la capçada i la quantitat de radiació "
        "interceptada. Intensitats de poda diferents, des de la poda severa fins a "
        "l'absència d'esporga, generen dossers amb volums i superfícies transpirants "
        "diferents. Com que la transpiració depèn de la superfície foliar activa i de "
        "la radiació interceptada, és previsible que els tractaments amb capçades més "
        "grans donin valors més alts d'ETa i que les podes més severes redueixin el "
        "consum d'aigua (Choné et al., 2001).")

replace("A més de l'efecte directe sobre la mida",
        "A més de l'efecte sobre la mida del dosser, el sistema de poda determina com "
        "es distribueixen la fusta i el fullatge i, per tant, com varia la intercepció "
        "de llum al llarg del dia. Aquesta variació diürna costa de capturar amb "
        "mesures puntuals, i és el motiu de combinar la teledetecció instantània amb "
        "dron amb la corba diürna de fIPAR descrita a l'apartat anterior.")

# ---------------------------------------------------------------------------
# 4. OBJECTIUS I HIPÒTESIS
# ---------------------------------------------------------------------------
replace("L'objectiu general d'aquest treball",
        "L'objectiu general és avaluar l'efecte dels tractaments de maneig de capçada "
        "sobre l'evapotranspiració de la vinya mitjançant teledetecció d'alta precisió "
        "amb dron i mesures de radiació PAR en camp.")

replace("Quantificar, mitjançant teledetecció",
        "Es tracta de quantificar l'ETa de cada tractament mantenint tots els "
        "tractaments en un mateix estat hídric, per comprovar si les capçades més "
        "reduïdes generen menys demanda hídrica, i de validar-ho amb mesures de camp "
        "(Ψstem i fIPAR). Si es confirma, el maneig de capçada permetria mantenir o "
        "millorar la qualitat del raïm gastant menys aigua.")

replace("Estimar les variables biofísiques del dosser",
        "Estimar les variables biofísiques del dosser (LAI, fIPAR diari, alçada i "
        "volum) per a cada tractament de poda a partir d'imatges multiespectrals i del "
        "núvol de punts fotogramètric obtinguts en quatre vols de dron al llarg del "
        "cicle vegetatiu.")

replace("Validar les estimacions de fIPAR",
        "Validar les estimacions de fIPAR derivades de les imatges de dron amb mesures "
        "simultànies de ceptòmetre i d'imatges hemisfèriques, seguint el protocol de "
        "Belaid et al. (2025).")

replace("H3. En mantenir-se un mateix estat hídric",
        "H3. Com que el reg individualitzat manté tots els tractaments en un mateix "
        "estat hídric, les diferències d'ETa i de CWSI responen a l'arquitectura del "
        "dosser i no a un estrès imposat diferent entre tractaments.")

# ---------------------------------------------------------------------------
# 5. MATERIALS I MÈTODES
# ---------------------------------------------------------------------------
replace("La Figura 4 resumeix el flux de treball",
        "La Figura 4 resumeix el flux de treball seguit en aquest TFM, des de "
        "l'adquisició de dades amb dron i en camp fins a l'anàlisi estadística de "
        "l'efecte dels tractaments, passant pel processament d'imatges, els models "
        "d'evapotranspiració i la validació. A continuació es descriu cada etapa.")

replace("La vinya, plantada l'any 2013",
        "La vinya, plantada l'any 2013, es condueix en doble cordó amb un marc de "
        "plantació d'1,7 m entre ceps i 2,5 m entre fileres. El reg s'aplica per "
        "degoteig seguint una estratègia de reg deficitari controlat (RDC), amb "
        "imposició del dèficit des del verolat fins a la verema. Cada tractament de "
        "maneig de capçada es rega de manera individualitzada per mantenir tots els "
        "tractaments en un mateix estat hídric, sigui quina sigui la mida de la "
        "capçada, de manera que les diferències observades es puguin atribuir a "
        "l'arquitectura del dosser. El potencial hídric de tija al migdia és "
        "l'indicador que es fa servir per controlar aquest estat hídric (Choné et al., "
        "2001).")

replace("L'assaig segueix un disseny en blocs",
        "L'assaig segueix un disseny en blocs amb cinc tractaments de maneig de "
        "capçada (esporga) i quatre blocs (I, II, III i IV), cosa que dona 20 "
        "parcel·les experimentals. Cada parcel·la experimental està formada per sis "
        "ceps consecutius d'una filera, excepte dues que en tenen cinc, i entre "
        "parcel·les es deixen dos ceps sense mostrejar per evitar els efectes de vora. "
        "Cada tractament reuneix, doncs, 24 ceps repartits entre els quatre blocs. El "
        "marc de plantació és d'1,7 m entre ceps i 2,5 m entre fileres. Els cinc "
        "tractaments es descriuen a la Taula 1 i la seva distribució a la parcel·la es "
        "representa a la Figura 5.")

replace("Es realitzaran quatre vols de dron",
        "Es faran quatre vols de dron al voltant del migdia solar, repartits al llarg "
        "del cicle vegetatiu, per seguir l'evolució estacional del dosser i de l'estat "
        "hídric. El dron porta dues càmeres: una de multiespectral de sis bandes "
        "(Altum-PT), amb la regió del red-edge, per calcular índexs de vegetació i "
        "estimar el LAI i la fIPAR; i una de tèrmica (FLIR) per obtenir la temperatura "
        "de la superfície. Els vols es planifiquen amb un solapament frontal i lateral "
        "elevat per garantir una bona reconstrucció fotogramètrica, i s'hi inclouen "
        "targets de calibratge radiomètric i punts de control terrestre "
        "georeferenciats amb GPS. El calendari de vols i de mesures de camp associades "
        "es resumeix a la Taula 2.")

replace("[A CONCRETAR] Model de dron i de cada càmera",
        "[A CONCRETAR] Model de dron; especificacions de l'Altum-PT i de la càmera "
        "FLIR (bandes espectrals, resolució tèrmica, GSD a l'altura de vol); altura de "
        "vol i solapaments; dates exactes dels quatre vols i quina fase fenològica "
        "cobreix cadascun (a la Taula 2 només n'hi ha tres definides); nombre i "
        "distribució dels punts de control; programari de planificació de vol.")

replace("Per separar les fonts del balanç energètic",
        "Per separar les fonts del balanç energètic es farà una segmentació "
        "supervisada de la coberta que distingirà els píxels purs de vegetació dels de "
        "sòl, i s'obtindran així la temperatura de dosser (Tc) i la de sòl (Ts) que "
        "necessiten els models TSEB i Shuttleworth–Wallace. Tota la cadena de "
        "processament segueix l'esquema descrit per Bellvert et al. (2021) per a "
        "cultius llenyosos amb imatges de dron.")

replace("A partir d'ambdues estimacions",
        "A partir d'aquestes dues estimacions, l'índex d'estrès hídric del cultiu es "
        "defineix com CWSI = 1 − ETa/ETp (Idso et al., 1981; Jackson et al., 1981), i "
        "pren valors propers a 0 sense estrès i propers a 1 en estrès màxim. El CWSI "
        "obtingut per teledetecció es contrasta amb les mesures de Ψstem per avaluar-ne "
        "la capacitat de diagnòstic (Bellvert et al., 2014). La Taula 3 resumeix les "
        "entrades i les sortides principals de cada model.")

replace("L'efecte dels tractaments de maneig de capçada",
        "L'efecte dels tractaments de maneig de capçada sobre les variables "
        "biofísiques del dosser (LAI, fIPAR, alçada i volum), sobre l'ETa, l'ETp i el "
        "CWSI, i sobre el Ψstem s'analitza comparant els cinc tractaments en cada data "
        "de mesura. Quan es compleixin els supòsits de normalitat dels residus i "
        "d'homogeneïtat de variàncies, s'aplicarà una anàlisi de la variància (ANOVA) "
        "amb separació de mitjanes pel test HSD de Tukey (p < 0,05), com s'ha fet en "
        "estudis equivalents en fruiters (Bellvert et al., 2021). Si no es compleixen, "
        "s'utilitzarà l'alternativa no paramètrica de Kruskal–Wallis, que és el criteri "
        "seguit en l'anàlisi de l'assaig ADAPTEX a l'IRTA. La concordança entre els "
        "mètodes de mesura de la fIPAR i entre el CWSI i el Ψstem s'avalua amb "
        "regressió lineal i estadístics de bondat d'ajust (R², RMSE i biaix).")

# ---------------------------------------------------------------------------
# 6. RESULTATS / DISCUSSIÓ
# ---------------------------------------------------------------------------
replace("Taules/figures de LAI, fIPAR, alçada",
        "Taules i figures de LAI, fIPAR, alçada i volum del dosser per als cinc "
        "tractaments (SE, EL, EM, ES i ES+DV) i per a cada data de vol; resultats del "
        "test (ANOVA amb Tukey o Kruskal–Wallis) amb lletres de separació de mitjanes.")

replace("CWSI per tractament; regressió CWSI",
        "CWSI per tractament; regressió CWSI–Ψstem; mapes d'estrès hídric. Com que el "
        "reg manté un mateix estat hídric entre tractaments, valoreu quina part del "
        "CWSI s'explica per l'arquitectura del dosser.")

doc.save(DST)
print("Fase 1 (text) completada ->", DST)
