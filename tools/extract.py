# -*- coding: utf-8 -*-
import json, datetime
import openpyxl
from openpyxl.utils import get_column_letter

wb = openpyxl.load_workbook('podcast_data.xlsx', data_only=True)

def td_sec(v):
    if v is None: return None
    if isinstance(v, datetime.timedelta): return int(v.total_seconds())
    if isinstance(v, datetime.time): return v.hour*3600 + v.minute*60 + v.second
    if isinstance(v, datetime.datetime): return int(v.hour*3600 + v.minute*60 + v.second)
    try: return int(float(v))
    except Exception: return None

def num(v):
    if v is None or v == '-' or v == 'nd': return None
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 4)
    except Exception:
        return None

def pct(v):
    v = num(v)
    return None if v is None else round(v, 4)

# ============ Suivi épisodes ============
ws = wb['Suivi épisodes']
episodes = []
for r in range(2, ws.max_row + 1):
    if ws.cell(row=r, column=1).value is None:
        continue
    g = lambda c: ws.cell(row=r, column=c).value
    d = g(8)
    date = d.strftime('%Y-%m-%d') if isinstance(d, datetime.datetime) else None
    scam_n, scam_r = g(27), g(28)
    episodes.append({
        'num': str(g(1)),
        'titre': g(2) or '',
        'accroche': g(3) if g(3) not in ('-',) else '',
        'genre': g(4) if g(4) not in ('-',) else '',
        'epoque': g(5) if g(5) not in ('-',) else '',
        'invites': g(6) if g(6) not in ('-',) else '',
        'dureeSec': td_sec(g(7)),
        'date': date,
        'statut': g(9) or '',
        'audioblog': num(g(10)),
        'youtube': num(g(11)),
        'lecturesVerifiees': num(g(12)),
        'spotify30': num(g(13)),
        'autre30': num(g(14)),
        'total30': num(g(15)),
        'totalAjuste30': num(g(16)),
        'auditeursUniques': num(g(17)),
        'ecoutesParAuditeur': pct(g(18)),
        'tempsLectureTotalSec': td_sec(g(19)),
        'tempsLectureMoySec': td_sec(g(20)),
        'pctLectureSpotify': pct(g(21)),
        'completion': pct(g(22)),
        'pctLectureYoutube': pct(g(23)),
        'retentionYt30': pct(g(24)),
        'ytEst30': pct(g(25)),
        'tauxConversion': pct(g(26)),
        'scamNatif': bool(scam_n) if isinstance(scam_n, bool) else None,
        'scamRadio': bool(scam_r) if isinstance(scam_r, bool) else None,
        'streams60': num(g(29)),
        'retentionYt60': pct(g(30)),
        'ytEst60': pct(g(31)),
        'streams60Total': num(g(32)),
    })

# ============ Plateformes ============
ws = wb['Plateformes']
platforms = []
for r in range(2, 9):
    name = ws.cell(row=r, column=1).value
    if not name or name == 'TOTAL':
        continue
    ea = ws.cell(row=r, column=3).value
    aj = ws.cell(row=r, column=8).value
    lec = num(ws.cell(row=r, column=2).value)
    # La ligne « Deezer » de l'onglet contient en réalité le relevé « Autre (hors Spotify) »
    # (909 = somme des 30s hors Spotify, cf. colonne N du suivi). Label corrigé.
    platforms.append({
        'nom': name,
        'lectures': lec if lec != 909 else None,
        'ecoutes30': None if ea in ('-', 'nd', None) or num(ea) == 909 else num(ea),
        'ecoutes30Ajustees': num(aj) if num(aj) != 909 else None,
    })
totals_plat = {
    'lectures': num(ws.cell(row=9, column=2).value),
    'ecoutes30': num(ws.cell(row=9, column=3).value),
    'ecoutes30Ajustees': num(ws.cell(row=9, column=8).value),
    'parEpisode': '121,4/épisode • 122,7 hors E0',
    'ratioYt30': pct(ws.cell(row=7, column=9).value),
}

# ============ TDB mensuel ============
ws = wb['TDB']
months = [ws.cell(row=14, column=c).value for c in range(3, 20)]  # C..S
def row_series(row, col_start=3, col_end=20):
    return [num(ws.cell(row=row, column=c).value) for c in range(col_start, col_end)]

tdb = {
    'mois': months,
    'episodes': row_series(15),
    'ecoutesSfc': row_series(16),
    'ecoutesParEp': [None if v is None else round(v, 1) for v in row_series(19)],
    'ecoutesNew': row_series(22),
    'ecoutesNewParEp': [None if v is None else round(v, 1) for v in row_series(25)],
    'conversionSpotify': row_series(28),
    'abonnesPlateformes': row_series(30),
    'abonnesReseaux': row_series(33),
    'ecoutesAjustees': row_series(36),
    'tempsLectureSec': [td_sec(ws.cell(row=33, column=c).value) for c in range(3, 20)],
    'tempsLectureMoySec': [td_sec(ws.cell(row=34, column=c).value) for c in range(3, 20)],
}
kpis_tdb = {
    'abonnesPlateformes': num(ws.cell(row=30, column=19).value),
    'abonnesInstagram': num(ws.cell(row=3, column=8).value),
    'abonnesFacebook': num(ws.cell(row=3, column=9).value),
    'joursDebut': num(ws.cell(row=6, column=11).value),
    'ecoutesParJour': pct(ws.cell(row=7, column=11).value),
    'ytdRatio30': pct(ws.cell(row=7, column=9).value),
    'ytdRatio60': pct(ws.cell(row=8, column=9).value),
    'episodeEnCours': ws.cell(row=2, column=13).value,
}

# ============ Découverte ============
ws = wb['Découverte']
dmonths = [ws.cell(row=3, column=c).value for c in range(2, 11)]  # B..J = Fév..Oct
def drow(row):
    return [num(ws.cell(row=row, column=c).value) for c in range(2, 11)]
discovery = {
    'mois': dmonths,
    'impressions': drow(4),
    'personnesTouchees': drow(5),
    'interessees': drow(6),
    'auditeurs': drow(7),
    'auditeursSpotify': drow(8),
    'dontRecurrents': drow(9),
    'dontNouveaux': drow(10),
    'tauxClic': drow(11),
    'tauxEcoute': drow(12),
    'tempsMoyenH': drow(13),
    'abonnement': drow(14),
    'conversion': drow(15),
    'lecturesSpotify': drow(16),
    'tauxConversionSpotify': drow(17),
    'tauxCompletionNouveaux': drow(18),
    'origineRecherche': drow(22),
    'origineAccueil': drow(23),
    'origineBibliotheque': drow(24),
    'partRecherche': drow(25),
    'partAccueil': drow(26),
    'partBibliotheque': drow(27),
    'episodesMois': drow(29),
    'episodesTotaux': drow(30),
    'dureeMoyH': drow(31),
}

# Tops épisodes par impressions : 3 périodes (Août 2026 R4-13, Sept 2026 R17-26, Oct 2026 R30-39)
tops = {}
for label, r0 in (('Août 2026', 4), ('Septembre 2026', 17), ('Octobre 2026', 30)):
    items = []
    for r in range(r0, r0 + 10):
        t = ws.cell(row=r, column=13).value
        if not t: continue
        d = ws.cell(row=r, column=18).value
        items.append({
            'titre': t,
            'impressions': num(ws.cell(row=r, column=14).value),
            'date': d.strftime('%Y-%m-%d') if isinstance(d, datetime.datetime) else None,
            'accueil': num(ws.cell(row=r, column=15).value),
            'recherche': num(ws.cell(row=r, column=16).value),
            'bibliotheque': num(ws.cell(row=r, column=17).value),
        })
    tops[label] = items

# ============ Checklist production (E25..E44) ============
tasks_map = {}
task_names = None
for name in wb.sheetnames:
    if not name.startswith('E'):
        continue
    ws = wb[name]
    local = []
    for r in range(2, ws.max_row + 1):
        t = ws.cell(row=r, column=1).value
        if not t: continue
        v = ws.cell(row=r, column=3).value
        dl = ws.cell(row=r, column=2).value
        local.append({
            'tache': t,
            'echeance': dl.strftime('%Y-%m-%d') if isinstance(dl, datetime.datetime) else None,
            'fait': bool(v) if v is not None else False,
        })
        if task_names is None:
            task_names = [x['tache'] for x in local]
    tasks_map[name] = {
        'taches': local,
        'faites': sum(1 for x in local if x['fait']),
        'total': len(local),
    }

# ============ Agrégats ============
def s(key, only=None):
    t = 0
    for e in episodes:
        if only and e['num'] not in only: continue
        v = e[key]
        if isinstance(v, (int, float)): t += v
    return round(t, 2)

publies = [e for e in episodes if e['statut'] == 'Publié']
encurs = [e for e in episodes if e['statut'] == 'En cours']

summary = {
    'episodesPublies': len(publies),
    'episodesEnCours': len(encurs),
    'episodesPlanifies': len(episodes) - len(publies) - len(encurs),
    'sommeSpotify30': s('spotify30'),
    'sommeAutre30': s('autre30'),
    'sommeTotal30': s('total30'),
    'sommeAjuste30': s('totalAjuste30'),
    'sommeAuditeurs': s('auditeursUniques'),
    'sommeLecturesVerifiees': s('lecturesVerifiees'),
    'sommeAudioblog': s('audioblog'),
    'sommeYoutube': s('youtube'),
    'sommeTempsSec': s('tempsLectureTotalSec'),
    'sommeStreams60': s('streams60'),
    'sommeStreams60Total': s('streams60Total'),
}

latest_pub = max(publies, key=lambda e: e['date'] or '')
data = {
    'meta': {
        'podcast': 'le beau rôle.',
        'genere': datetime.date.today().isoformat(),
        'dernierEpEpPublie': latest_pub['num'],
        'derniereDatePublie': latest_pub['date'],
        'source': 'UCr6usYb-Suivi épisodes v2.xlsx',
    },
    'summary': summary,
    'episodes': episodes,
    'platforms': platforms,
    'totalsPlatforms': totals_plat,
    'tdb': tdb,
    'kpisTdb': kpis_tdb,
    'discovery': discovery,
    'topsImpressions': tops,
    'production': tasks_map,
    'taskNames': task_names,
}

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

print('OK —', len(episodes), 'épisodes,',
      len(tdb['mois']), 'mois TDB,', len(discovery['mois']), 'mois découverte,',
      sum(len(v) for v in tops.values()), 'tops,',
      len(tasks_map), 'checklists')
print('Spotify30:', summary['sommeSpotify30'], '| Autre30:', summary['sommeAutre30'],
      '| Total30:', summary['sommeTotal30'], '| Ajusté:', summary['sommeAjuste30'])
print('Plateformes: écoutes30', totals_plat['ecoutes30'], '| ajustées', totals_plat['ecoutes30Ajustees'])
print('Taille data.json:', len(open('data.json', encoding='utf-8').read())//1024, 'Ko')
