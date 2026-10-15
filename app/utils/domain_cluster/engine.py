"""Deterministic, explainable domain planning. No network or AI required."""
from __future__ import annotations

import csv
import hashlib
import io
import re
import random
from collections import Counter
from datetime import datetime, timezone
from difflib import SequenceMatcher
from functools import lru_cache
from urllib.parse import urlsplit

import tldextract

RULE_VERSION = '2026.10.1'
MAX_DOMAINS = 5000
MAX_INPUT = 1_000_000
# Bundled PSL: parsing never downloads a suffix list or sends input off-server.
EXTRACT = tldextract.TLDExtract(suffix_list_urls=(), cache_dir=None, include_psl_private_domains=True)
FAMILIES = {
    'outreach': ('Outreach & Prospecting', 'outreach prospect prospecting engage engagement connect connection contact discovery relationship rapport dialogue converse conversation introduce'),
    'sales': ('Sales & Pipeline', 'sales selling pipeline revenue commercial deals opportunity opportunities lead leads demand businessdevelopment leadgeneration'),
    'growth': ('Agency & Growth', 'agency agencies growth scaling scale expand expansion development consultancy consultant advisory advise accelerator accelerate'),
    'partners': ('Partnerships & Connections', 'partner partners partnership partnerships alliance connector network bridge collaboration introductions collaborate allies'),
    'atlantic': ('UK–US & Transatlantic', 'uk usa us american british atlantic transatlantic crossborder marketentry'),
    'international': ('International Expansion', 'global international overseas worldwide export expansion internationalbusiness marketdevelopment'),
    'acquisition': ('Acquisition & Dealflow', 'acquisition acquire acquisitions dealflow mergers takeover buyout'),
    # Distinct commercial intents, rather than artificial subdivisions of full families.
    'strategy': ('Business Strategy', 'strategy strategic foresight roadmap planning direction'),
    'research': ('Market Research', 'research insights intelligence survey analytics evidence'),
    'customer': ('Client Success', 'customer clientsuccess retention loyalty service experience'),
    'talent': ('Talent & Leadership', 'talent leadership recruitment executive staffing'),
}
NEUTRAL = set('north south east west summit cedar oak silver blue clear bright prime field stone harbor harbour horizon crest willow maple anchor beacon sterling compass vista venture business group company team firm hub collective point peak path house works studio and the solutions capital markets market entry new fresh first next spring river green open bold wise true fair'.split())
STATUSES = {'Active', 'Paused', 'Planned', 'Retired'}
DEFAULT_SETTINGS = {'capacity': 10, 'watch': 7, 'near': 9, 'sensitivity': .72, 'dictionary': {}, 'themes': [], 'industry': '', 'geography': '', 'default_import': 'add'}


def now():
    return datetime.now(timezone.utc).isoformat()


def normalize(value: str) -> str:
    value = value.strip().strip('"\'<>[](){},;.!?')
    if not value or any(c.isspace() for c in value):
        raise ValueError('Enter a domain, mailbox or URL.')
    if value.lower().startswith('mailto:'):
        value = value[7:]
    if '://' not in value and '@' in value:
        if value.count('@') != 1 or not value.split('@')[0]:
            raise ValueError('Malformed mailbox address.')
        value = value.rsplit('@', 1)[1]
    parsed = urlsplit(value if '://' in value else 'https://' + value)
    if parsed.scheme.lower() not in {'http', 'https'} or parsed.username or parsed.password:
        raise ValueError('Unsupported URL or credentials in URL.')
    try:
        parsed.port  # Validate malformed ports.
        host = (parsed.hostname or '').rstrip('.').encode('idna').decode('ascii').lower()
    except (ValueError, UnicodeError):
        raise ValueError('Invalid hostname or port.') from None
    if len(host) > 253 or not all(re.fullmatch(r'[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?', p) for p in host.split('.')):
        raise ValueError('Invalid hostname.')
    result = EXTRACT(host)
    if not result.suffix or not result.domain:
        raise ValueError('No recognised public suffix or registrable domain.')
    return result.top_domain_under_public_suffix


def parse(text: str) -> dict:
    if len(text.encode('utf-8')) > MAX_INPUT:
        raise ValueError('Input limit is 1 MB. Import in smaller batches.')
    seen, invalid, duplicates = set(), [], 0
    # CSV and spreadsheet separators; retain each original line for immediate feedback only.
    for line_no, line in enumerate(text.splitlines(), 1):
        if not line.strip():
            continue
        try:
            cells = next(csv.reader([line]))
        except csv.Error:
            cells = [line]
        for cell in cells:
            for token in re.split(r'[\s;]+', cell.strip()):
                if not token:
                    continue
                try:
                    domain = normalize(token)
                except ValueError as exc:
                    invalid.append({'line': line_no, 'input': token[:300], 'original_line': line[:500], 'reason': str(exc)})
                    continue
                if domain in seen:
                    duplicates += 1
                seen.add(domain)
                if len(seen) > MAX_DOMAINS:
                    raise ValueError(f'Limit is {MAX_DOMAINS:,} unique domains per portfolio.')
    return {'domains': sorted(seen), 'unique': len(seen), 'duplicates': duplicates, 'invalid': invalid}


def settings(raw=None):
    result = {**DEFAULT_SETTINGS, **(raw or {})}
    if not 1 <= result['capacity'] <= 10:
        raise ValueError('Capacity must be between 1 and 10.')
    if not 1 <= result['watch'] <= result['near'] <= result['capacity']:
        raise ValueError('Use warning ≤ near-limit ≤ capacity.')
    return result


def dictionary(config):
    terms = {k: set(v[1].split()) for k, v in FAMILIES.items()}
    for family, words in config.get('dictionary', {}).items():
        if family not in terms:
            raise ValueError('Choose an existing semantic family for custom terms.')
        terms[family].update(words)
    return terms


@lru_cache(maxsize=8192)
def segment(label: str, vocabulary: tuple[str, ...]) -> tuple[str, ...]:
    """Maximum coverage segmentation with a cost for unknown characters.

    Short geographic tokens are accepted only when the entire label is segmented;
    this prevents 'us' in 'business' from becoming transatlantic evidence.
    """
    known = set(vocabulary)
    n = len(label)
    scores, paths = [-float('inf')] * (n + 1), [[] for _ in range(n + 1)]
    scores[0] = 0
    for i in range(n):
        if scores[i] - 1 > scores[i + 1]:
            scores[i + 1], paths[i + 1] = scores[i] - 1, paths[i] + [label[i]]
        for size in range(2, min(n - i, 25) + 1):
            word = label[i:i + size]
            if word in known:
                score = scores[i] + size * 2 + size * size * .01
                if score > scores[i + size]:
                    scores[i + size], paths[i + size] = score, paths[i] + [word]
    path = paths[n]
    unknown = any(len(w) == 1 for w in path)
    return tuple(w for w in path if len(w) > 1 and (len(w) > 2 or not unknown))


def classify(domain: str, config=None) -> dict:
    config = settings(config)
    terms = dictionary(config)
    label = EXTRACT(domain).domain
    vocabulary = tuple(sorted(set.union(NEUTRAL, *terms.values())))
    tokens = set(segment(label.replace('-', ''), vocabulary))
    evidence, scores = {}, {}
    for family, words in terms.items():
        matches = tokens & words
        fuzzy = []
        # Fuzzy matching is conservative and limited to the whole unknown label.
        if not tokens:
            fuzzy = [w for w in words if len(w) >= 6 and SequenceMatcher(None, label, w).ratio() >= .9]
        if matches or fuzzy:
            evidence[family] = sorted(matches) + [f'{w} (near spelling)' for w in sorted(fuzzy)]
            scores[family] = sum(1.2 if len(w) >= 5 else 1 for w in matches) + len(fuzzy) * .75
    ranked = sorted(scores, key=lambda k: (-scores[k], list(FAMILIES).index(k)))
    primary = ranked[0] if ranked else 'review'
    confidence = .92 if ranked and (len(ranked) == 1 or scores[ranked[0]] - scores[ranked[1]] >= .5) else .6 if ranked else .15
    review = confidence < config['sensitivity']
    reason = ('Matched ' + ', '.join(evidence[primary]) + ' to ' + FAMILIES[primary][0] + '.') if ranked else 'No strong commercial-intent evidence; manual classification needed.'
    if len(ranked) > 1:
        reason += ' Potential semantic overlap: ' + ', '.join(FAMILIES[k][0] for k in ranked[1:]) + '.'
    return {'automatic_cluster': primary, 'cluster': primary, 'tags': ranked[1:], 'semantic_families': ranked, 'confidence': confidence, 'review_required': review, 'reason': reason, 'tokens': sorted(tokens), 'rule_version': RULE_VERSION}


def risk(count, config):
    if count > config['capacity']:
        return 'Over Limit'
    if count == config['capacity']:
        return 'Full'
    if count >= config['near']:
        return 'Near Limit'
    if count >= config['watch']:
        return 'Watch'
    return 'Healthy'


def empty_state():
    return {'domains': [], 'clusters': {}, 'settings': dict(DEFAULT_SETTINGS), 'recommendations': [], 'snapshots': [], 'history': []}


def analyse(state: dict) -> dict:
    config = settings(state.get('settings'))
    clusters = {k: {'id': k, 'name': v[0], 'domains': [], 'active': 0, 'planned': 0} for k, v in FAMILIES.items()}
    clusters['review'] = {'id': 'review', 'name': 'Review Required', 'domains': [], 'active': 0, 'planned': 0}
    for key, custom in state.get('clusters', {}).items():
        clusters[key] = {'id': key, 'name': custom['name'], 'domains': [], 'active': 0, 'planned': 0}
    output = []
    for item in sorted(state.get('domains', []), key=lambda r: r['domain']):
        features = classify(item['domain'], config)
        record = {**item, **features}
        if item.get('override'):
            record['cluster'] = item['override']
            record['reason'] += ' User-confirmed assignment.'
        if item.get('reviewed') or item.get('override'):
            record['review_required'] = False
        cluster = record['cluster']
        if cluster not in clusters:
            raise ValueError('Unknown cluster assignment.')
        clusters[cluster]['domains'].append(record)
        if record['status'] == 'Active':
            clusters[cluster]['active'] += 1
        elif record['status'] == 'Planned':
            clusters[cluster]['planned'] += 1
        output.append(record)
    used = []
    for cluster in clusters.values():
        projected = cluster['active'] + cluster['planned']
        cluster.update(projected=projected, capacity=config['capacity'], risk=risk(cluster['active'], config), projected_risk=risk(projected, config), remaining=max(0, config['capacity'] - projected))
        if cluster['id'] == 'review':
            cluster.update(risk='Review Required', projected_risk='Review Required', remaining=0)
        if cluster['domains'] or cluster['id'] in state.get('clusters', {}):
            used.append(cluster)
    # Evidence counts remain intact even after manual reassignment; recommendations
    # cannot evade a full natural family by creating a different display cluster.
    exposure = Counter()
    for r in output:
        if r['status'] in {'Active', 'Planned'}:
            exposure.update(r['semantic_families'])
    summary = {'active': sum(r['status'] == 'Active' for r in output), 'planned': sum(r['status'] == 'Planned' for r in output), 'clusters': sum(c['active'] > 0 and c['id'] != 'review' for c in used), 'watch': sum(c['risk'] in {'Watch', 'Near Limit'} for c in used), 'full': sum(c['risk'] == 'Full' for c in used), 'over': sum(c['risk'] == 'Over Limit' for c in used), 'remaining': sum(c['remaining'] for c in used if c['id'] != 'review'), 'review': sum(r['review_required'] for r in output if r['status'] in {'Active', 'Planned'})}
    warnings = [{'cluster': c['name'], 'id': c['id'], 'count': c['active'], 'projected': c['projected'], 'risk': c['projected_risk'], 'kind': 'capacity'} for c in used if c['projected_risk'] not in {'Healthy', 'Review Required'}]
    for family, count in sorted(exposure.items()):
        assigned = next((c['projected'] for c in used if c['id'] == family), 0)
        if count > assigned and count >= config['watch']:
            warnings.append({'cluster': FAMILIES[family][0], 'id': family, 'count': count, 'projected': count, 'risk': risk(count, config), 'kind': 'semantic_overlap'})
    return {'domains': output, 'clusters': used, 'summary': summary, 'warnings': warnings, 'semantic_exposure': dict(exposure), 'settings': config, 'rule_version': RULE_VERSION}


PREFIXES = 'cedar summit sterling north willow bright anchor beacon maple clear horizon silver oak compass prime vista'.split()
PATTERNS = {
    'outreach': [('outreach',), ('rapport',), ('prospect',)],
    'sales': [('revenue',), ('pipeline',), ('commercial',)],
    'growth': [('agency',), ('advisory',), ('growth',)],
    'partners': [('alliance',), ('partners',), ('bridge',)],
    'atlantic': [('atlantic',), ('transatlantic',)],
    'international': [('export',), ('global',)],
    'acquisition': [('acquisition',), ('dealflow',)],
    'strategy': [('strategy',), ('foresight',)],
    'research': [('insights',), ('research',)],
    'customer': [('retention',), ('loyalty',)],
    'talent': [('leadership',), ('talent',)],
}


def validate_candidate(candidate, state, accepted=()):
    words = candidate.get('words', [])
    domain = candidate.get('domain', '')
    if not isinstance(words, list) or not 2 <= len(words) <= 3 or not all(isinstance(w, str) and re.fullmatch('[a-z]+', w) and w in set.union(NEUTRAL, *(set(v[1].split()) for v in FAMILIES.values())) for w in words):
        return False
    if domain != ''.join(words) + '.com' or len(domain) > 67:
        return False
    existing = {r['domain'] for r in state['domains']}
    if domain in existing or domain in {r['domain'] for r in accepted}:
        return False
    result = analyse(state)
    feature = classify(domain, state['settings'])
    if feature['review_required'] or feature['cluster'] == 'review':
        return False
    exposure = Counter(result['semantic_exposure'])
    assigned = Counter({c['id']: c['projected'] for c in result['clusters']})
    for r in accepted:
        f = classify(r['domain'], state['settings'])
        exposure.update(f['semantic_families'])
        assigned[f['cluster']] += 1
    maximum = result['settings']['capacity']
    return assigned[feature['cluster']] < maximum and all(exposure[f] < maximum for f in feature['semantic_families'])


def generate(state, batch=0, excluded=()):
    result = analyse(state)
    accepted = []
    existing = {r['domain'] for r in state['domains']} | set(excluded)
    # Stable seeded rotation permits new batches while preserving deterministic runs.
    seed = int(hashlib.sha256(('\n'.join(sorted(existing)) + str(batch)).encode()).hexdigest()[:8], 16)
    candidates = [(family, [prefix, *suffix]) for prefix in PREFIXES for family, patterns in PATTERNS.items() for suffix in patterns]
    random.Random(seed).shuffle(candidates)
    themes = state['settings'].get('themes', [])
    if themes:
        candidates.sort(key=lambda c: c[0] not in themes)
    # Cache counts once; candidate loops are linear rather than portfolio all-pairs.
    exposure = Counter(result['semantic_exposure'])
    assigned = Counter({c['id']: c['projected'] for c in result['clusters']})
    original = assigned.copy()
    capacity = result['settings']['capacity']
    for _, words in candidates:
        domain = ''.join(words) + '.com'
        if domain in existing:
            continue
        feature = classify(domain, state['settings'])
        family = feature['cluster']
        if feature['review_required'] or family == 'review' or assigned[family] >= capacity or any(exposure[f] >= capacity for f in feature['semantic_families']):
            continue
        assigned[family] += 1
        exposure.update(feature['semantic_families'])
        accepted.append({'domain': domain, 'words': words, 'cluster': family, 'cluster_name': next((c['name'] for c in result['clusters'] if c['id'] == family), FAMILIES[family][0]), 'current': original[family], 'projected': assigned[family], 'rationale': f'{words[0].title()} paired with {" ".join(words[1:])} conveys {FAMILIES[family][0].lower()}.', 'availability': 'Availability not checked', 'checked_at': None})
        existing.add(domain)
        if len(accepted) == 10:
            break
    if len(accepted) != 10:
        return {'recommendations': [], 'shortfall': 10 - len(accepted), 'qualifying_count': len(accepted), 'message': 'Fewer than ten candidates meet the current capacity policy. Review capacity or expand the naming dictionary; no partial batch was approved.'}
    for r in accepted:
        r['projected'] = assigned[r['cluster']]
    return {'recommendations': accepted, 'shortfall': 0}


def import_preview(state, parsed, mode='add'):
    if mode not in {'add', 'replace'}:
        raise ValueError('Choose add or replace import.')
    incoming = set(parsed['domains'])
    existing = {r['domain'] for r in state['domains']}
    active = {r['domain'] for r in state['domains'] if r['status'] == 'Active'}
    return {'added': sorted(incoming - existing), 'existing': sorted(incoming & existing), 'missing': sorted(active - incoming), 'reactivated': sorted(incoming & (existing - active)) if mode == 'replace' else [], 'mode': mode, 'invalid_count': len(parsed['invalid']), 'duplicates': parsed['duplicates']}


def apply_import(state, parsed, mode):
    import copy
    state = copy.deepcopy(state)
    preview = import_preview(state, parsed, mode)
    stamp = now()
    incoming = set(parsed['domains'])
    for row in state['domains']:
        if mode == 'replace':
            if row['domain'] in incoming:
                row['status'] = 'Active'
            elif row['status'] == 'Active':
                row['status'] = 'Paused'
    state['domains'].extend({'domain': d, 'status': 'Active', 'added_at': stamp, 'owner': '', 'platform': '', 'notes': '', 'reviewed': False, 'override': None} for d in preview['added'])
    if len(state['domains']) > MAX_DOMAINS:
        raise ValueError(f'Portfolio limit is {MAX_DOMAINS:,} domains, including historical entries.')
    return state, preview
