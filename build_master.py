import pandas as pd, numpy as np, json
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
m=pd.read_pickle('master.pkl');imgs=json.load(open('dl/imgs.json'))
m['Photo']=np.where(m['Code article'].isin(imgs.keys()),'Oui','A fournir')
issues=[]
for _,r in m[m.Afficher=='O'].iterrows():
    if r.Photo!='Oui': issues.append((r['Code article'],r.Nom,'Photo manquante'))
    if r.Bloc=='PRECO': issues.append((r['Code article'],r.Nom,'Prix, EAN, contenance, colisage et visuel à fournir'))
issues+= [('10906523','Les Florales - Sans Sulfites','Domaine saisi « Domiane Demiane » dans Vinistoria : faute de frappe probable'),
 ('24600221','Château Carbonnieux rouge CB','Domaine affiché « Les Capucins » ; EAN à préfixe 0 (bouteille export ?) ; stock 3 bt'),
 ('13703020','Le Daron -Tout','Libellé Vinistoria « Le Daron -Tout » à nettoyer ?'),
 ('35800225','Chablis Brocard','Stock 0 au 30/09 : affiché « Réappro en cours »'),
 ('13700222','Closerie La Mahèle 2022','Stock 3 bt ; le 2023 (13700223, 402 bt) porte probablement le même EAN : basculer sur le 2023 ?'),
 ('722999FT','Chiroulet','Article à créer dans Vinistoria, masqué en attendant'),
 ('FEZAS-BR','Le Bouquet de Fruits Rouges','Gencod bouteille toujours à confirmer. Le mail Fezas du 29/10 corrige le seul Grand Bouquet (3439902359041). Pour le Bouquet Rouge : 3439902369040 (fiche Les Gouttes) ou 3439902319045 (déduit du code caisse 3439902319144)'),
 ('FEZAS-GB','Le Grand Bouquet','Gencod 3439902359041 confirmé par Fezas le 29/10 (remplace 3439902369040) ; il ne correspond plus au code caisse 3439902369149 de la fiche : code caisse à faire confirmer'),
 ('BROCARD-BM','Belle Margote','Nom à confirmer : « Belle Margote » (étiquette, fiche, mail du 23/10) mais « Grande Margote » dans le mail du 26/10 ; millésime non communiqué (étiquette 2020)'),
 ('BROCARD-*','Logistique Brocard','MAD par palette de 600 bt à Transvins non confirmée (le domaine dit « idéalement non panaché », la proposition VINOM est 300 + 300) ; passera en Dispo 24 h une fois la palette reçue'),
 ('CSM-*','Château de Saint-Martin / Bleutiful','Commission : question posée (7 % AOP/IGP ?), sans réponse ; colisage des magnums à demander'),
 ('CMGC-*','Compagnie Médocaine des Grands Crus','EN ATTENTE (consigne du 29/10) : 11 vins masqués, à ne pas traiter tant que les gencods et visuels ne sont pas reçus ; seule la liste Excel compte, la brochure « Exclusivités 2026 tarif G » est écartée ; commission à préciser'),
 ('LORENTZ-PG-MO','Pinot Gris Médaille d’Or 2023 (Jules Muller)','Visuel reçu : étiquette « Réserve » 2023 avec pastille Or Gilbert & Gaillard 2024, rattaché à la Médaille d’Or (le Pinot Gris Réserve 3328776861197 est masqué). Vérifier que ce gencod correspond bien à la bouteille avec pastille'),
 ('LORENTZ-*','Jules Muller (Gustave Lorentz)','4 références affichées (celles avec visuel), 8 masquées faute de visuel ; colisage 6 déduit de la palette (77 cartons de 6) ; aucun des 12 gencods n’existe dans Vinistoria : articles à créer'),
 ('MALDANT-*','J&L Maldant : franco en euros','Franco 500 € HT d’achat (colonne « Franco (€ HT) ») : jauge en euros à ajouter au modèle lors de la régénération ; d’ici là affiché en texte'),
 ('12100624 / 12102223','Maldant en Dispo 24 h','Doublons : les références TVD en stock (Aligoté 2024, 12 bt ; Chardonnay 2023, 54 bt) remplacent l’offre direct domaine, masquée. Gencods J&L Maldant à saisir dans Vinistoria ; vérifier l’étiquette des bouteilles en stock (le Chardonnay est libellé « Maldant Pauvelot »)'),
 ('MALDANT-*','J&L Maldant : commission','Commission VINOM non précisée'),
 ('LORENTZ-*','Jules Muller : franco','Règle « prix franco de port impérativement » : le tarif Lorentz est franco Transvins (120 cols), pas franco magasin. Faire confirmer un prix franco livré magasin'),
 ('13702624','Côté Océan blanc 2024 (TVD)','Doublon : la référence en stock Transvins (58 bt, T2 3,90 €) remplace l’offre direct domaine Tutiac (2025, 2,93 à 3,35 €), masquée. Gencod 3412911034161 à saisir sur l’article 2024'),
 ('TUTIAC-*','Les Vignerons de Tutiac : références écartées','Belordre blanc 2023 (au tarif, sans gencod ni visuel) et BIB 3 L Côté Océan blanc et rosé (gencods fournis, absents du tarif, sans visuel) : masqués'),
 ('20600424 / 20600121','Vannières en Dispo 24 h','Doublons : les références TVD en stock (Bandol blanc 2024, 58 bt, T2 17,30 € ; Bandol rouge 2021, 55 bt, T2 18,85 €) remplacent l’offre direct domaine (15,50 €), masquée. Gencods de la gamme GD à contrôler sur une bouteille en stock avant saisie dans Vinistoria : le rouge en stock est un 2021 (visuel 2022) et Vinistoria donnait 3302901131072 au rouge 2019'),
 ('VANNIERES-*','Château Vannières : millésimes','Millésimes non indiqués dans le mail, lus sur les visuels : rosés Bandol et Côtes de Provence 2026, alors que la récolte 2026 vient d’être vendangée : à confirmer (2025 probablement livré d’ici le printemps)'),
 ('VANNIERES-HERITAGE-RS','Héritage Vannières rosé','Appellation et millésime non précisés (étiquette « La Cadière d’Azur ») : à demander'),
 ('VANNIERES-*','Château Vannières : colisage','Colisage non communiqué : 6 bt supposé'),
 ('VEDRENNE-SIN-*','Maison Vedrenne : gamme Sîn','Gamme conseillée par Vedrenne : 11 gencods sur 15, millésimes non précisés'),
 ('VEDRENNE-* / VED-*','Maison Vedrenne : gencods','Gencods reçus le 29/10 (fichier « Nouveaux Bordeaux » pour Sîn, tarif gencods 2022 pour le reste) : 45 références avec gencod (dernier millésime au tarif 2026), toutes masquées faute de visuel HD. Sans gencod : Sîn Malbec, Sîn Bordeaux rouge, Sîn Pessac blanc, Sîn Sauternes et 28 autres lignes du tarif (voir Vedrenne_tarif_selection_20261029.xlsx)'),
 ('VED-LES-BELLES-VIGNES / -BIB','Les Belles Vignes 75 cl et BIB','Deux gencods côte à côte sur le tarif 2022 : 3499148895635 attribué à la 75 cl, 3499147372045 au BIB, par position : à confirmer'),
 ('VED-MARTILLAC / VED-PERLE-DARCHE','Liquoreux Vedrenne','Gencods attribués par position sur la page (Martillac 3525150175189, Perle d’Arche 3760138254467) : à confirmer'),
 ('VED-*','Maison Vedrenne : gencods 2022','Gencods du tarif 2022 appliqués aux millésimes 2026 : à contrôler sur bouteille (un gencod peut changer avec l’étiquette). Colisage : cartons de 12 au tarif 2026, de 6 au tarif 2022'),
 ('VEDRENNE-*','Maison Vedrenne : conditions','En-tête « FRANCO » mais conditions anglaises « ex cellar - Libourne » : faire confirmer le franco magasin. Cartons de 12 bt (commande au carton de 12), palette NIMP15 facturée 13 € si non échangée, contre-étiquette 0,15 €/bt'),
 ('PORTIA-*','Bodega Portia (via Vedrenne)','Roble 2023 6,90 €, 10 Meses 2022 8,40 €, Crianza 2021 10,30 € : prix non franco (« +/10 » : 10 % ou 0,10 € ?), sans gencod ni visuel : non intégrés'),
 ('VL-*','Veyret Latour : périmètre','Tarif GMS 2026 « petits châteaux », onglet « sans Anim » (prix net HT franco 120 bt) : 71 références, 24 affichées (gencod valide + visuel), 42 masquées faute de visuel, 5 masquées pour gencod invalide ou en double'),
 ('VL-TOUTIGEAC / VL-HAUT-CORMIER','Gencod en double','Même gencod 3334582211880 pour Toutigeac 2023 et Haut Cormier 2017 ; le visuel reçu porte « Grand Cormier », pas « Haut Cormier »'),
 ('VL-FERREAU-BELAIR / VL-GRAVETTES','Gencod en double','Même gencod 3463410001208 pour Ferreau Belair et Les Gravettes'),
 ('VL-VIEUX-MOULIN','Château Vieux Moulin','Gencod « 063443837914 » à 12 chiffres : invalide'),
 ('VL-SAINT-CLAIR / VL-VIEUX-MOULIN-ROQUES / VL-SAINT-GREGOIRE','Gencods étrangers','Préfixes 0634438 (États-Unis) et 742 (Honduras) : clés valides mais à contrôler sur la bouteille'),
 ('VL-*','Veyret Latour : prix et commission','Prix affiché = « Prix Net » ; le « Prix Brut » est supérieur de 5 % (net / 0,95) : commission VINOM incluse dans l’écart ? Onglet « avec Anim » (prix plus élevés, animations magasin) non utilisé'),
 ('VL-*','Veyret Latour : millésimes des visuels','Visuels d’un autre millésime que le tarif : Bédouret, Bellegrave (2017), Tour Fonrazade (2020), Hauts de Martillac (2018), Graveyron et Pingoy (2007)'),
 ('VDP-*','Vision de Paix (Via Caritatis)','EN ATTENTE : aucun gencod fourni (tarif et fiches 2025). Rouge et rosé avec visuel, blanc sans visuel ni fiche. Franco Île-de-France 36/60/120 bt : 5,85 / 5,20 / 4,80 € HT ; PVC 7,49 € TTC ; commission non précisée'),
 ('TRIBAUT-*','Champagne Reine Dhuicq (Tribaut)','5 cuvées avec gencod et visuel, affichées en direct domaine. Cartons de 6 ; franco dès 16 cartons = 96 bt ; options d’animation (« dynamique ») dès 30 cartons ou box de 192 bt non intégrées ; commission non précisée'),
 ('(relevé 02/11)','Relevé EAN GDD2 du 29/10','21 références TVD ajoutées et affichées ; Domaine Casanova blanc (242003NM) validé par le scan ; 12 visuels DirectCave ajoutés (dont 3 d’un autre millésime). Stock au 30/09 : export STOCK à rafraîchir'),
 ('35800225','Chablis Brocard (TVD)','Visuel reçu le 05/11 avec l’étiquette « Famille Brocard », comme la gamme GD (BROCARD-CH) ; le gencod scanné (3436802500016) diffère de celui de la gamme GD (3436801911011) : vérifier que l’étiquette en stock est bien celle-ci'),
 ('FIERCE-MILD / FIERCE-MILD-LAGER','Fierce Mild','Prix catalogue 2,32 € HT/bt (55,60 € HT le carton de 24 x 33 cl, PVC 3,49 € TTC). Achat VINOM 39,50 € HT/carton rendu Paris, palette complète minimum, paiement d’avance. Gencods à fournir ; format bouteille retenu (existe aussi en canette)'),
 ('336001NM / 252005NM / 252001NM','Champagnes scannés le 05/11','Gencods scannés différents de Vinistoria : Baudry Blanc de Noirs 3760089510070 (Vinistoria 3760089510063), Gosset Brut Extra 3353211019994 (3353210019995), Gosset Grande Réserve 3353211012216 (3353211012254) : à corriger dans Vinistoria. Baudry sans visuel'),
 ('336001NM','Baudry Blanc de Noirs Héritage','Visuel provisoire : étiquette « Distinction Brut » de Baudry, pas « Blanc de Noirs Héritage »'),
 ('(tous)','Colonne Visible','À remplir : O pour chaque bouteille à l’étiquette accrocheuse')]
wb=Workbook();H=Font(name='Arial',bold=True,color='FFFFFF');F=PatternFill('solid',fgColor='1F4E79');A=Font(name='Arial',size=10)
INP=PatternFill('solid',fgColor='FFF2CC')
ws=wb.active;ws.title='Lisez-moi'
T=['Catalogue GMS VINOM : fichier maître (28/10/2026)',
'Ce fichier sert de source quand le catalogue est régénéré : il n’est à charger nulle part. Le modifier puis le renvoyer à Claude, qui renvoie un nouvel index.html à déposer sur GitHub.',
'Colonnes à remplir (fond jaune) : « Afficher » (O = visible dans le catalogue, N = masqué), « Visible » (O = étiquette accrocheuse : badge « Visible en rayon », filtre dédié et affichage en tête), « Labels » (Bio, Demeter, Sans sulfites ajoutés, HVE, Vegan, séparés par des virgules).',
'Labels pré-remplis depuis Vinistoria (champs Label) et DirectCave (Certification) : à compléter ou corriger.',
'Bloc : TVD (Dispo 24 h, Transvins), DD (Direct domaine), PRECO (Précommande). Prix = T2 HT à la bouteille. Stock = « Stock Dispo » de l’export Vinistoria du 30/09/2026 (indicateur : Dispo 24 h ≥ 24 bt, Stock limité 1 à 23 bt, Réappro en cours à 0).',
'Les 41 références dont l’EAN est déjà dans Vinistoria sont présentes mais masquées (N) : à passer en O si l’étiquette convient. Onglet « A completer » : ce qu’il manque.']
for i,t in enumerate(T,1):
    c=ws.cell(i,1,t);c.font=Font(name='Arial',bold=(i==1),size=12 if i==1 else 10);c.alignment=Alignment(wrap_text=True)
ws.column_dimensions['A'].width=140
cols=['Afficher','Visible','Labels','Bloc','Catégorie','Code article','Nom','Domaine / Marque','Appellation','Millésime','Couleur','Région','Contenance','Colisage (bt/carton)','EAN','Prix HT bouteille','Paliers (cols:prix)','Fournisseur','Franco (cols)','Franco (€ HT)','Conditions domaine','Commission VINOM','Stock dispo (bt)','Photo','Image (URL)','Fiche DirectCave','Descriptif court','Note interne']
W={'Afficher':9,'Visible':8,'Labels':22,'Bloc':7,'Catégorie':14,'Code article':12,'Nom':32,'Domaine / Marque':24,'Appellation':26,'Millésime':9,'Couleur':10,'Région':14,'Contenance':10,'Colisage (bt/carton)':9,'EAN':16,'Prix HT bouteille':10,'Paliers (cols:prix)':18,'Fournisseur':22,'Franco (cols)':9,'Franco (€ HT)':9,'Conditions domaine':40,'Commission VINOM':22,'Stock dispo (bt)':9,'Photo':9,'Image (URL)':30,'Fiche DirectCave':30,'Descriptif court':50,'Note interne':45}
def put(name,df,widths,inputs=()):
    w=wb.create_sheet(name);w.append(list(df.columns))
    for c in w[1]: c.font=H;c.fill=F;c.alignment=Alignment(wrap_text=True,vertical='center')
    for row in df.itertuples(index=False): w.append([None if (isinstance(v,float) and np.isnan(v)) or v=='' else v for v in row])
    for row in w.iter_rows(min_row=2):
        for c in row: c.font=A;c.alignment=Alignment(vertical='top')
    for i,col in enumerate(df.columns,1):
        L=get_column_letter(i);w.column_dimensions[L].width=widths.get(col,14)
        if col in ('EAN','Code article'):
            for c in w[L][1:]: c.number_format='@'
        if col=='Prix HT bouteille':
            for c in w[L][1:]: c.number_format='#,##0.00 €'
        if col in inputs:
            for c in w[L][1:]: c.fill=INP
    w.row_dimensions[1].height=32;w.freeze_panes='H2';w.auto_filter.ref=w.dimensions
    return w
mm=m.sort_values(['Bloc','Afficher','Région','Nom'],ascending=[False,False,True,True])[cols]
w=put('Produits',mm,W,('Afficher','Visible','Labels','Paliers (cols:prix)','Franco (cols)','Franco (€ HT)'))
dv=DataValidation(type='list',formula1='"O,N"',allow_blank=True);w.add_data_validation(dv);n=len(mm)+1;dv.add(f'A2:A{n}');dv.add(f'B2:B{n}')
put('A completer',pd.DataFrame(issues,columns=['Code article','Produit','À compléter / à vérifier']),{'Code article':12,'Produit':34,'À compléter / à vérifier':90})
wb.save('/mnt/user-data/outputs/catalogue-gms/Catalogue_GD_master_20261028.xlsx');print('ok',len(mm))
