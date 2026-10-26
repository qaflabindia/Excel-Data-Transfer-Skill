"""Packaging/source checks only; this is not an Excel/VBA compiler."""
from pathlib import Path
import re
root = Path(__file__).resolve().parents[1]
modules = list((root / 'vba').glob('*.bas'))
assert len(modules) == 2
for path in modules:
    text = path.read_text()
    assert 'Option Explicit' in text
    assert not re.search(r'On Error Resume Next', text, re.I)
    assert max(map(len, text.splitlines())) < 1024
    starts = re.findall(r'^(?:Public |Private )?(Sub|Function)\s+(\w+)', text, re.M | re.I)
    ends = re.findall(r'^End (Sub|Function)\b', text, re.M | re.I)
    assert len(starts) == len(ends), (path, len(starts), len(ends))
    names = [n.lower() for _, n in starts]
    assert len(names) == len(set(names)), path
    print(f'{path.name}: {len(starts)} procedures, structural source checks passed')
production = (root / 'vba/Transfer.bas').read_text()
assert not re.search(r'\.(Unprotect|Protect)\b', production, re.I)
assert not re.search(r'\b(ActiveSheet|Selection)\b|\.Select\b', production, re.I)
for name in ['DiscoverDataRegions', 'DetectHeaderBands', 'GetCellHeaderKeys', 'ScopedFind',
             'BuildDryRun', 'ExecuteMapping', 'VerifyPreservation', 'Reconcile']:
    assert re.search(r'(Sub|Function) ' + name + r'\b', production)
assert 'If w.ProtectContents Then' in production
assert 'Set candidate = cell.CurrentRegion' in production
print('Production safety-pattern checks passed; Excel compilation/execution still required.')
