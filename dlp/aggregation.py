def aggregate(findings, policy):
    specific = max((f['score'] for f in findings if f.get('family') == 'cellveilxl'), default=0)
    families = {}
    for finding in findings:
        family = finding.get('family')
        if family != 'cellveilxl':
            families[family] = max(families.get(family, 0), finding['score'])
    strongest = max(families.values(), default=0)
    corroboration = sum(score >= 35 for score in families.values())
    generic = min(policy.generic_cap, strongest + min(20, policy.corroboration_bonus * max(0, corroboration - 1)))
    return {'formatting_steganography': {'risk_score': generic, 'signal_families': sorted(families)},
            'cellveilxl': {'risk_score': specific, 'generic_support_score': generic}}
