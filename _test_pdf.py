import sys, os
sys.path.insert(0, r'c:\Users\Aria\seadrive_root\Maximili\Meine Bibliotheken\UKE\STX-Zahlen-Projekt\Dashboard-MG_2026\scripts')
from parse_pdf_reports import parse_treat_par_pdf

folder = r'C:\Users\Aria\Desktop\testPDF'
for fname in sorted(os.listdir(folder)):
    if not fname.endswith('.pdf'):
        continue
    p = parse_treat_par_pdf(os.path.join(folder, fname))
    if not p:
        print(fname + ': FEHLER')
        continue
    ver = p.get('app_version', '?')
    print('=== ' + fname + ' v' + str(ver) + ' ===')
    print('  plan=' + str(p.get('plan_name')) + '  fractions=' + str(p.get('fractions')) + '  n_ptv=' + str(len(p.get('ptv_details', []))))
    for ptv in p.get('ptv_details', []):
        print('  PTV: ' + str(ptv))
    print()
