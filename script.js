// ─── CATÁLOGO ───
const CAT = {
  "PLANCHAS LAMINADAS EN FRIO (LAF)": [
  {codigo:"LAF-0.3", descripcion:"1/84 (0.3) × 3 × 8 pies", peso_unit:5.12, unidad:"pies"},
  {codigo:"LAF-0.4", descripcion:"1/64 (0.4) × 3 × 8 pies", peso_unit:6.82, unidad:"pies"},
  {codigo:"LAF-0.5", descripcion:"1/54 (0.5) × 3 × 8 pies", peso_unit:8.53, unidad:"pies"},
  {codigo:"LAF-0.6", descripcion:"1/40 (0.6) × 4 × 8 pies", peso_unit:13.56, unidad:"pies"},
  {codigo:"LAF-0.8", descripcion:"1/32 (0.8) × 4 × 8 pies", peso_unit:18.09, unidad:"pies"},
  {codigo:"LAF-0.9", descripcion:"1/27 (0.9) × 4 × 8 pies", peso_unit:20.35, unidad:"pies"},
  {codigo:"LAF-1.0", descripcion:"1/24 (1.0) × 4 × 8 pies", peso_unit:22.61, unidad:"pies"},
  {codigo:"LAF-1.2", descripcion:"1/20 (1.2) × 4 × 8 pies", peso_unit:27.13, unidad:"pies"},
  {codigo:"LAF-1.5", descripcion:"1/16 (1.5) × 4 × 8 pies", peso_unit:33.91, unidad:"pies"},
  {codigo:"LAF-2.0", descripcion:"5/64 (2.0) × 4 × 8 pies", peso_unit:45.22, unidad:"pies"}
],
  "PLANCHAS DELGADAS LAMINADAS EN CALIENTE (LAC)": [
  {codigo:"LAC-2.0", descripcion:"Plancha LAC 2.0 × 1200 × 2400 mm", peso_unit:45.22, unidad:"mm"},
  {codigo:"LAC-2.4", descripcion:"Plancha LAC 2.4 × 1200 × 2400 mm", peso_unit:54.26, unidad:"mm"},
  {codigo:"LAC-2.5", descripcion:"Plancha LAC 2.5 × 1200 × 2400 mm", peso_unit:56.52, unidad:"mm"},
  {codigo:"LAC-3.0", descripcion:"Plancha LAC 3.0 × 1200 × 2400 mm", peso_unit:67.82, unidad:"mm"},
  {codigo:"LAC-4.0", descripcion:"Plancha LAC 4.0 × 1200 × 2400 mm", peso_unit:90.43, unidad:"mm"},
  {codigo:"LAC-4.5", descripcion:"Plancha LAC 4.5 × 1200 × 2400 mm", peso_unit:101.74, unidad:"mm"}
],

  "PLANCHAS GRUESAS LAMINADAS EN CALIENTE (LAC)": [
  {codigo:"LAC-6.0",  descripcion:"Plancha LAC 6.0  × 1200 × 2400 mm", peso_unit: 135.65,  unidad:"mm"},
  {codigo:"LAC-8.0",  descripcion:"Plancha LAC 8.0  × 1200 × 2400 mm", peso_unit: 180.86,  unidad:"mm"},
  {codigo:"LAC-9.0",  descripcion:"Plancha LAC 9.0  × 1200 × 2400 mm", peso_unit: 203.47,  unidad:"mm"},
  {codigo:"LAC-12.0", descripcion:"Plancha LAC 12.0 × 1200 × 2400 mm", peso_unit: 271.30,  unidad:"mm"},
  {codigo:"LAC-16.0", descripcion:"Plancha LAC 16.0 × 1200 × 2400 mm", peso_unit: 361.73,  unidad:"mm"},
  {codigo:"LAC-20.0", descripcion:"Plancha LAC 20.0 × 1200 × 2400 mm", peso_unit: 452.16,  unidad:"mm"},
  {codigo:"LAC-25.0", descripcion:"Plancha LAC 25.0 × 1200 × 2400 mm", peso_unit: 565.20,  unidad:"mm"},
  {codigo:"LAC-32.0", descripcion:"Plancha LAC 32.0 × 1200 × 2400 mm", peso_unit: 723.46,  unidad:"mm"},
  {codigo:"LAC-38.0", descripcion:"Plancha LAC 38.0 × 1200 × 2400 mm", peso_unit: 859.10,  unidad:"mm"},
  {codigo:"LAC-50.0", descripcion:"Plancha LAC 50.0 × 1200 × 2400 mm", peso_unit: 130.40,  unidad:"mm"},
],

"PLANCHAS ESTRIADAS (LAGRIMA/ROMBO) ASTM A-569": [
  {codigo:"EST-3/32-3.28x8", descripcion:"Plancha Estriada 3/32 × 3.28' × 8 pies", peso_unit:  50.15, unidad:"pies"},
  {codigo:"EST-3/32-4x8",    descripcion:"Plancha Estriada 3/32 × 4 × 8 pies",     peso_unit:  60.14, unidad:"pies"},
  {codigo:"EST-1/8-3.28x8",  descripcion:"Plancha Estriada 1/8 × 3.28' × 8 pies",  peso_unit:  59.53, unidad:"pies"},
  {codigo:"EST-1/8-4x8",     descripcion:"Plancha Estriada 1/8 × 4 × 8 pies",      peso_unit:  71.44, unidad:"pies"},
  {codigo:"EST-3/16-4x8",    descripcion:"Plancha Estriada 3/16 × 4 × 8 pies",     peso_unit: 105.35, unidad:"pies"},
  {codigo:"EST-1/4-4x8",     descripcion:"Plancha Estriada 1/4 × 4 × 8 pies",      peso_unit: 139.27, unidad:"pies"},
],

  "PLANCHAS ASTM A-36": [
  {codigo:"A36-1/8-5x20",    descripcion:"Plancha A36 1/8 × 5 × 20 pies (3.0 × 1500 × 6000 mm)",    peso_unit: 211.95,  unidad:"mm"},
  {codigo:"A36-3/16-5x20",   descripcion:"Plancha A36 3/16 × 5 × 20 pies (4.5 × 1500 × 6000 mm)",   peso_unit: 317.93,  unidad:"mm"},
  {codigo:"A36-1/4-5x20",    descripcion:"Plancha A36 1/4 × 5 × 20 pies (6.0 × 1500 × 6000 mm)",    peso_unit: 423.90,  unidad:"mm"},
  {codigo:"A36-5/16-5x20",   descripcion:"Plancha A36 5/16 × 5 × 20 pies (8.0 × 1500 × 6000 mm)",   peso_unit: 565.20,  unidad:"mm"},
  {codigo:"A36-3/8-5x20",    descripcion:"Plancha A36 3/8 × 5 × 20 pies (9.0 × 1500 × 6000 mm)",    peso_unit: 635.85,  unidad:"mm"},
  {codigo:"A36-1/2-5x20",    descripcion:"Plancha A36 1/2 × 5 × 20 pies (12.0 × 1500 × 6000 mm)",   peso_unit: 847.80,  unidad:"mm"},
  {codigo:"A36-5/8-5x20",    descripcion:"Plancha A36 5/8 × 5 × 20 pies (16.0 × 1500 × 6000 mm)",   peso_unit: 1130.40, unidad:"mm"},
  {codigo:"A36-3/4-5x20",    descripcion:"Plancha A36 3/4 × 5 × 20 pies (20.0 × 1500 × 6000 mm)",   peso_unit: 1413.00, unidad:"mm"},
  {codigo:"A36-1-5x20",      descripcion:"Plancha A36 1\" × 5 × 20 pies (25.0 × 1500 × 6000 mm)",   peso_unit: 1766.25, unidad:"mm"},
  {codigo:"A36-1-1/4-5x20",  descripcion:"Plancha A36 1 1/4 × 5 × 20 pies (32.0 × 1500 × 6000 mm)",peso_unit: 2260.80, unidad:"mm"},
  {codigo:"A36-1-1/2-5x20",  descripcion:"Plancha A36 1 1/2 × 5 × 20 pies (38.0 × 1500 × 6000 mm)",peso_unit: 2684.70, unidad:"mm"},
  {codigo:"A36-2-5x20",      descripcion:"Plancha A36 2 × 5 × 20 pies (50.0 × 1500 × 6000 mm)",     peso_unit: 3532.50, unidad:"mm"},
  {codigo:"A36-3-5x20",      descripcion:"Plancha A36 3 × 5 × 20 pies (75.0 × 1500 × 6000 mm)",     peso_unit: 5298.75, unidad:"mm"},
  {codigo:"A36-4-5x20",      descripcion:"Plancha A36 4 × 5 × 20 pies (100.0 × 1500 × 6000 mm)",    peso_unit: 7065.00, unidad:"mm"},
  {codigo:"A36-5/16-8x20",   descripcion:"Plancha A36 5/16 × 8 × 20 pies (8.0 × 2400 × 6000 mm)",   peso_unit: 904.32,  unidad:"mm"},
  {codigo:"A36-3/8-8x20",    descripcion:"Plancha A36 3/8 × 8 × 20 pies (9.0 × 2400 × 6000 mm)",    peso_unit: 1017.36, unidad:"mm"},
  {codigo:"A36-1/2-8x20",    descripcion:"Plancha A36 1/2 × 8 × 20 pies (12.0 × 2400 × 6000 mm)",   peso_unit: 1356.48, unidad:"mm"},
  {codigo:"A36-5/8-8x20",    descripcion:"Plancha A36 5/8 × 8 × 20 pies (16.0 × 2400 × 6000 mm)",   peso_unit: 1808.64, unidad:"mm"},
  {codigo:"A36-3/4-8x20",    descripcion:"Plancha A36 3/4 × 8 × 20 pies (20.0 × 2400 × 6000 mm)",   peso_unit: 2260.80, unidad:"mm"},
  {codigo:"A36-1-8x20",      descripcion:"Plancha A36 1\" × 8 × 20 pies (25.0 × 2400 × 6000 mm)",   peso_unit: 2826.00, unidad:"mm"},
  {codigo:"A36-1-1/4-8x20",  descripcion:"Plancha A36 1 1/4 × 8 × 20 pies (32.0 × 2400 × 6000 mm)",peso_unit: 3617.28, unidad:"mm"},
  {codigo:"A36-1-1/2-8x20",  descripcion:"Plancha A36 1 1/2 × 8 × 20 pies (38.0 × 2400 × 6000 mm)",peso_unit: 4295.52, unidad:"mm"},
  {codigo:"A36-2-8x20",      descripcion:"Plancha A36 2 × 8 × 20 pies (50.0 × 2400 × 6000 mm)",     peso_unit: 5652.00, unidad:"mm"},
],
"PLANCHAS ASTM A-283 GRADO C": [
  {codigo:"A283-1/8-5x20",   descripcion:"Plancha A283-C 1/8 × 5 × 20 pies (3.0 × 1500 × 6000 mm)",   peso_unit: 211.95,  unidad:"mm"},
  {codigo:"A283-3/16-5x20",  descripcion:"Plancha A283-C 3/16 × 5 × 20 pies (4.5 × 1500 × 6000 mm)",  peso_unit: 317.93,  unidad:"mm"},
  {codigo:"A283-1/4-5x20",   descripcion:"Plancha A283-C 1/4 × 5 × 20 pies (6.0 × 1500 × 6000 mm)",   peso_unit: 423.90,  unidad:"mm"},
  {codigo:"A283-5/16-5x20",  descripcion:"Plancha A283-C 5/16 × 5 × 20 pies (8.0 × 1500 × 6000 mm)",  peso_unit: 565.20,  unidad:"mm"},
  {codigo:"A283-3/8-5x20",   descripcion:"Plancha A283-C 3/8 × 5 × 20 pies (9.0 × 1500 × 6000 mm)",   peso_unit: 635.85,  unidad:"mm"},
  {codigo:"A283-1/2-5x20",   descripcion:"Plancha A283-C 1/2 × 5 × 20 pies (12.0 × 1500 × 6000 mm)",  peso_unit: 847.80,  unidad:"mm"},
  {codigo:"A283-5/8-5x20",   descripcion:"Plancha A283-C 5/8 × 5 × 20 pies (16.0 × 1500 × 6000 mm)",  peso_unit: 1130.40, unidad:"mm"},
  {codigo:"A283-3/4-5x20",   descripcion:"Plancha A283-C 3/4 × 5 × 20 pies (20.0 × 1500 × 6000 mm)",  peso_unit: 1413.00, unidad:"mm"},
  {codigo:"A283-1-5x20",     descripcion:"Plancha A283-C 1\" × 5 × 20 pies (25.0 × 1500 × 6000 mm)",  peso_unit: 1766.25, unidad:"mm"},
  {codigo:"A283-1-1/4-5x20", descripcion:"Plancha A283-C 1 1/4 × 5 × 20 pies (32.0 × 1500 × 6000 mm)",peso_unit: 2260.80, unidad:"mm"},
  {codigo:"A283-1-1/2-5x20", descripcion:"Plancha A283-C 1 1/2 × 5 × 20 pies (38.0 × 1500 × 6000 mm)",peso_unit: 2684.70, unidad:"mm"},
  {codigo:"A283-2-5x20",     descripcion:"Plancha A283-C 2 × 5 × 20 pies (50.0 × 1500 × 6000 mm)",    peso_unit: 3532.50, unidad:"mm"},
],

"PLANCHAS COR-TEN ASTM A-242": [
  {codigo:"A242-1/8-5x20",  descripcion:"Plancha Cor-Ten A242 1/8 × 5 × 20 pies",   peso_unit: 224.31, unidad:"pies"},
  {codigo:"A242-3/16-5x20", descripcion:"Plancha Cor-Ten A242 3/16 × 5 × 20 pies",  peso_unit: 336.47, unidad:"pies"},
  {codigo:"A242-1/4-5x20",  descripcion:"Plancha Cor-Ten A242 1/4 × 5 × 20 pies",   peso_unit: 448.63, unidad:"pies"},
  {codigo:"A242-5/16-5x20", descripcion:"Plancha Cor-Ten A242 5/16 × 5 × 20 pies",  peso_unit: 565.20, unidad:"pies"},
  {codigo:"A242-3/8-5x20",  descripcion:"Plancha Cor-Ten A242 3/8 × 5 × 20 pies",   peso_unit: 672.94, unidad:"pies"},
  {codigo:"A242-1/2-5x20",  descripcion:"Plancha Cor-Ten A242 1/2 × 5 × 20 pies",   peso_unit: 897.26, unidad:"pies"},
],

"PLANCHAS ANTIABRASIVAS 400 Y 500 BHN": [
  {codigo:"ANTIA-1/4-5x20",  descripcion:"Plancha Antiabrasiva 400/500 BHN 1/4 × 5 × 20 pies",   peso_unit:  461.88, unidad:"pies"},
  {codigo:"ANTIA-5/16-5x20", descripcion:"Plancha Antiabrasiva 400/500 BHN 5/16 × 5 × 20 pies",  peso_unit:  581.90, unidad:"pies"},
  {codigo:"ANTIA-3/8-5x20",  descripcion:"Plancha Antiabrasiva 400/500 BHN 3/8 × 5 × 20 pies",   peso_unit:  692.82, unidad:"pies"},
  {codigo:"ANTIA-1/2-5x20",  descripcion:"Plancha Antiabrasiva 400/500 BHN 1/2 × 5 × 20 pies",   peso_unit:  923.77, unidad:"pies"},
  {codigo:"ANTIA-5/8-5x20",  descripcion:"Plancha Antiabrasiva 400/500 BHN 5/8 × 5 × 20 pies",   peso_unit: 1154.71, unidad:"pies"},
  {codigo:"ANTIA-3/4-5x20",  descripcion:"Plancha Antiabrasiva 400/500 BHN 3/4 × 5 × 20 pies",   peso_unit: 1385.65, unidad:"pies"},
  {codigo:"ANTIA-1-5x20",    descripcion:"Plancha Antiabrasiva 400/500 BHN 1\" × 5 × 20 pies",   peso_unit: 1847.53, unidad:"pies"},
],

"PLANCHAS NAVALES ASTM 131 GRADO A / LRS GRADO A": [
  {codigo:"NAV-1/8-5x20",    descripcion:"Plancha Naval 1/8 × 5 × 20 pies (3.0 × 1500 × 6000 mm)",    peso_unit:  224.31, unidad:"mm"},
  {codigo:"NAV-3/16-6x20",   descripcion:"Plancha Naval 3/16 × 6 × 20 pies (4.5 × 1800 × 6000 mm)",   peso_unit:  403.72, unidad:"mm"},
  {codigo:"NAV-1/4-6x20",    descripcion:"Plancha Naval 1/4 × 6 × 20 pies (6.4 × 1800 × 6000 mm)",    peso_unit:  538.35, unidad:"mm"},
  {codigo:"NAV-5/16-6x20",   descripcion:"Plancha Naval 5/16 × 6 × 20 pies (8.0 × 1800 × 6000 mm)",   peso_unit:  678.24, unidad:"mm"},
  {codigo:"NAV-3/8-6x20",    descripcion:"Plancha Naval 3/8 × 6 × 20 pies (9.5 × 1800 × 6000 mm)",    peso_unit:  807.46, unidad:"mm"},
  {codigo:"NAV-1/2-6x20",    descripcion:"Plancha Naval 1/2 × 6 × 20 pies (12.5 × 1800 × 6000 mm)",   peso_unit: 1076.71, unidad:"mm"},
  {codigo:"NAV-5/8-6x20",    descripcion:"Plancha Naval 5/8 × 6 × 20 pies (16.0 × 1800 × 6000 mm)",   peso_unit: 1345.88, unidad:"mm"},
  {codigo:"NAV-3/4-6x20",    descripcion:"Plancha Naval 3/4 × 6 × 20 pies (20.0 × 1800 × 6000 mm)",   peso_unit: 1615.06, unidad:"mm"},
  {codigo:"NAV-1-6x20",      descripcion:"Plancha Naval 1\" × 6 × 20 pies (25.0 × 1800 × 6000 mm)",   peso_unit: 2153.41, unidad:"mm"},
  {codigo:"NAV-1-1/4-6x20",  descripcion:"Plancha Naval 1 1/4 × 6 × 20 pies (32.0 × 1800 × 6000 mm)",peso_unit: 2691.77, unidad:"mm"},
  {codigo:"NAV-1-1/2-6x20",  descripcion:"Plancha Naval 1 1/2 × 6 × 20 pies (38.0 × 1800 × 6000 mm)",peso_unit: 3230.19, unidad:"mm"},
  {codigo:"NAV-1/4-8x20",    descripcion:"Plancha Naval 1/4 × 8 × 20 pies (6.4 × 2400 × 6000 mm)",    peso_unit:  740.04, unidad:"mm"},
  {codigo:"NAV-5/16-8x20",   descripcion:"Plancha Naval 5/16 × 8 × 20 pies (8.0 × 2400 × 6000 mm)",   peso_unit:  925.04, unidad:"mm"},
  {codigo:"NAV-3/8-8x20",    descripcion:"Plancha Naval 3/8 × 8 × 20 pies (9.5 × 2400 × 6000 mm)",    peso_unit: 1098.49, unidad:"mm"},
  {codigo:"NAV-1/2-8x20",    descripcion:"Plancha Naval 1/2 × 8 × 20 pies (12.5 × 2400 × 6000 mm)",   peso_unit: 1445.38, unidad:"mm"},
  {codigo:"NAV-5/8-8x20",    descripcion:"Plancha Naval 5/8 × 8 × 20 pies (16.0 × 2400 × 6000 mm)",   peso_unit: 1850.09, unidad:"mm"},
  {codigo:"NAV-3/4-8x20",    descripcion:"Plancha Naval 3/4 × 8 × 20 pies (20.0 × 2400 × 6000 mm)",   peso_unit: 2312.61, unidad:"mm"},
  {codigo:"NAV-1-8x20",      descripcion:"Plancha Naval 1\" × 8 × 20 pies (25.0 × 2400 × 6000 mm)",   peso_unit: 2890.76, unidad:"mm"},
  {codigo:"NAV-1-1/4-8x20",  descripcion:"Plancha Naval 1 1/4 × 8 × 20 pies (32.0 × 2400 × 6000 mm)",peso_unit: 3700.18, unidad:"mm"},
  {codigo:"NAV-1-1/2-8x20",  descripcion:"Plancha Naval 1 1/2 × 8 × 20 pies (38.0 × 2400 × 6000 mm)",peso_unit: 4393.96, unidad:"mm"},
],

"PLANCHAS GALVANIZADAS LISAS ASTM A-653 / JIS G-3302": [
  {codigo:"GALV-0.25", descripcion:"Plancha Galv. Lisa 0.25 × 920 × 1830 mm (1/100 × 3 × 6 pies)",   peso_unit:  3.61, unidad:"pies"},
  {codigo:"GALV-0.30", descripcion:"Plancha Galv. Lisa 0.30 × 920 × 1830 mm (1/80 × 3 × 6 pies)",    peso_unit:  4.27, unidad:"pies"},
  {codigo:"GALV-0.40", descripcion:"Plancha Galv. Lisa 0.40 × 920 × 1830 mm (1/64 × 3 × 6 pies)",    peso_unit:  5.59, unidad:"pies"},
  {codigo:"GALV-0.50", descripcion:"Plancha Galv. Lisa 0.50 × 920 × 1830 mm (1/54 × 3 × 6 pies)",    peso_unit:  6.91, unidad:"pies"},
  {codigo:"GALV-0.60", descripcion:"Plancha Galv. Lisa 0.60 × 1200 × 2400 mm (1/40 × 4 × 8 pies)",   peso_unit: 14.08, unidad:"pies"},
  {codigo:"GALV-0.90", descripcion:"Plancha Galv. Lisa 0.90 × 1200 × 2400 mm (1/27 × 4 × 8 pies)",   peso_unit: 20.87, unidad:"pies"},
  {codigo:"GALV-1.00", descripcion:"Plancha Galv. Lisa 1.00 × 1200 × 2400 mm (1/24 × 4 × 8 pies)",   peso_unit: 23.13, unidad:"pies"},
  {codigo:"GALV-1.20", descripcion:"Plancha Galv. Lisa 1.20 × 1200 × 2400 mm (1/20 × 4 × 8 pies)",   peso_unit: 27.65, unidad:"pies"},
  {codigo:"GALV-1.50", descripcion:"Plancha Galv. Lisa 1.50 × 1200 × 2400 mm (1/16 × 4 × 8 pies)",   peso_unit: 34.43, unidad:"pies"},
],

"PLANCHAS ZINCADAS ONDULADAS (CALAMINAS) ASTM 361 / JIS G-3302": [
  {codigo:"CALM-0.20-850x1800",  descripcion:"Calamina Zincada 0.20 × 850 × 1800 mm",   peso_unit:  3.01, unidad:"mm"},
  {codigo:"CALM-0.22-657x1800",  descripcion:"Calamina Zincada 0.22 × 657 × 1800 mm",   peso_unit:  2.61, unidad:"mm"},
  {codigo:"CALM-0.22-830x1800",  descripcion:"Calamina Zincada 0.22 × 830 × 1800 mm",   peso_unit:  3.21, unidad:"mm"},
  {codigo:"CALM-0.25-830x1800",  descripcion:"Calamina Zincada 0.25 × 830 × 1800 mm",   peso_unit:  3.61, unidad:"mm"},
  {codigo:"CALM-0.30-830x1830",  descripcion:"Calamina Zincada 0.30 × 830 × 1830 mm",   peso_unit:  4.27, unidad:"mm"},
  {codigo:"CALM-0.40-830x1830",  descripcion:"Calamina Zincada 0.40 × 830 × 1830 mm",   peso_unit:  5.59, unidad:"mm"},
  {codigo:"CALM-0.50-830x1830",  descripcion:"Calamina Zincada 0.50 × 830 × 1830 mm",   peso_unit:  6.91, unidad:"mm"},
  {codigo:"CALM-0.60-830x1830",  descripcion:"Calamina Zincada 0.60 × 830 × 1830 mm",   peso_unit:  8.23, unidad:"mm"},
  {codigo:"CALM-0.60-830x2400",  descripcion:"Calamina Zincada 0.60 × 830 × 2400 mm",   peso_unit: 10.80, unidad:"mm"},
  {codigo:"CALM-0.60-830x3600",  descripcion:"Calamina Zincada 0.60 × 830 × 3600 mm",   peso_unit: 16.20, unidad:"mm"},
],

"PLANCHAS ACERO INOXIDABLE AISI 304 Y 316": [
  {codigo:"INOX-LAF-5/64-4x8",   descripcion:"Plancha Inox LAF 5/64 × 4 × 8 pies (2.0 mm)",    peso_unit:  47.13, unidad:"plancha"},
  {codigo:"INOX-LAF-3/32-4x8",   descripcion:"Plancha Inox LAF 3/32 × 4 × 8 pies (2.5 mm)",    peso_unit:  58.92, unidad:"plancha"},
  {codigo:"INOX-LAF-1/8-4x8",    descripcion:"Plancha Inox LAF 1/8 × 4 × 8 pies (3.0 mm)",     peso_unit:  70.70, unidad:"plancha"},
  {codigo:"INOX-LAC-1/8-4x8",    descripcion:"Plancha Inox LAC 1/8 × 4 × 8 pies (3.0 mm)",     peso_unit:  70.70, unidad:"plancha"},
  {codigo:"INOX-LAC-1/8-5x10",   descripcion:"Plancha Inox LAC 1/8 × 5 × 10 pies (3.0 mm)",    peso_unit: 110.51, unidad:"plancha"},
  {codigo:"INOX-LAC-3/16-4x8",   descripcion:"Plancha Inox LAC 3/16 × 4 × 8 pies (4.5 mm)",    peso_unit: 106.05, unidad:"plancha"},
  {codigo:"INOX-LAC-3/16-5x10",  descripcion:"Plancha Inox LAC 3/16 × 5 × 10 pies (4.5 mm)",   peso_unit: 165.76, unidad:"plancha"},
  {codigo:"INOX-LAC-1/4-4x8",    descripcion:"Plancha Inox LAC 1/4 × 4 × 8 pies (6.0 mm)",     peso_unit: 141.40, unidad:"plancha"},
  {codigo:"INOX-LAC-1/4-5x10",   descripcion:"Plancha Inox LAC 1/4 × 5 × 10 pies (6.0 mm)",    peso_unit: 221.02, unidad:"plancha"},
  {codigo:"INOX-LAC-3/8-4x8",    descripcion:"Plancha Inox LAC 3/8 × 4 × 8 pies (9.0 mm)",     peso_unit: 212.11, unidad:"plancha"},
  {codigo:"INOX-LAC-3/8-5x10",   descripcion:"Plancha Inox LAC 3/8 × 5 × 10 pies (9.0 mm)",    peso_unit: 331.52, unidad:"plancha"},
  {codigo:"INOX-LAC-1/2-4x8",    descripcion:"Plancha Inox LAC 1/2 × 4 × 8 pies (12.7 mm)",    peso_unit: 299.31, unidad:"plancha"},
  {codigo:"INOX-LAC-1/2-5x10",   descripcion:"Plancha Inox LAC 1/2 × 5 × 10 pies (12.7 mm)",   peso_unit: 467.82, unidad:"plancha"},
],

"ANGULOS ALTA RESISTENCIA ASTM A-572 GRADO 50": [
  {codigo:"ANG572-2x2x3/16",    descripcion:"Ángulo A572-50 2\" × 2\" × 3/16\"",      peso_unit: 21.79, unidad:"barra"},
  {codigo:"ANG572-2x2x1/4",     descripcion:"Ángulo A572-50 2\" × 2\" × 1/4\"",       peso_unit: 28.49, unidad:"barra"},
  {codigo:"ANG572-2.5x2.5x3/16",descripcion:"Ángulo A572-50 2 1/2\" × 2 1/2\" × 3/16\"", peso_unit: 27.42, unidad:"barra"},
  {codigo:"ANG572-2.5x2.5x1/4", descripcion:"Ángulo A572-50 2 1/2\" × 2 1/2\" × 1/4\"",  peso_unit: 36.61, unidad:"barra"},
  {codigo:"ANG572-3x3x1/4",     descripcion:"Ángulo A572-50 3\" × 3\" × 1/4\"",       peso_unit: 43.75, unidad:"barra"},
  {codigo:"ANG572-3x3x5/16",    descripcion:"Ángulo A572-50 3\" × 3\" × 5/16\"",      peso_unit: 54.47, unidad:"barra"},
],
"ANGULOS ESTRUCTURALES ASTM A-36": [
  {codigo:"ANG36-1.5x1.5x3/32", descripcion:"Ángulo A36 1 1/2\" × 1 1/2\" × 3/32\"", peso_unit:  8.292, unidad:"barra"},
  {codigo:"ANG36-1.5x1.5x1/8",  descripcion:"Ángulo A36 1 1/2\" × 1 1/2\" × 1/8\"",  peso_unit: 10.983, unidad:"barra"},
  {codigo:"ANG36-1.5x1.5x3/16", descripcion:"Ángulo A36 1 1/2\" × 1 1/2\" × 3/16\"", peso_unit: 16.072, unidad:"barra"},
  {codigo:"ANG36-1.5x1.5x1/4",  descripcion:"Ángulo A36 1 1/2\" × 1 1/2\" × 1/4\"",  peso_unit: 20.894, unidad:"barra"},
  {codigo:"ANG36-1.75x1.75x1/8",  descripcion:"Ángulo A36 1 3/4\" × 1 3/4\" × 1/8\"",  peso_unit: 12.858, unidad:"barra"},
  {codigo:"ANG36-1.75x1.75x3/16", descripcion:"Ángulo A36 1 3/4\" × 1 3/4\" × 3/16\"", peso_unit: 18.929, unidad:"barra"},
  {codigo:"ANG36-1.75x1.75x1/4",  descripcion:"Ángulo A36 1 3/4\" × 1 3/4\" × 1/4\"",  peso_unit: 24.733, unidad:"barra"},
  {codigo:"ANG36-2x2x1/8",   descripcion:"Ángulo A36 2\" × 2\" × 1/8\"",   peso_unit: 14.733, unidad:"barra"},
  {codigo:"ANG36-2x2x3/16",  descripcion:"Ángulo A36 2\" × 2\" × 3/16\"",  peso_unit: 21.787, unidad:"barra"},
  {codigo:"ANG36-2x2x1/4",   descripcion:"Ángulo A36 2\" × 2\" × 1/4\"",   peso_unit: 28.483, unidad:"barra"},
  {codigo:"ANG36-2x2x5/16",  descripcion:"Ángulo A36 2\" × 2\" × 5/16\"",  peso_unit: 35.002, unidad:"barra"},
  {codigo:"ANG36-2x2x3/8",   descripcion:"Ángulo A36 2\" × 2\" × 3/8\"",   peso_unit: 41.966, unidad:"barra"},
  {codigo:"ANG36-2.5x2.5x3/16", descripcion:"Ángulo A36 2 1/2\" × 2 1/2\" × 3/16\"", peso_unit: 27.412, unidad:"barra"},
  {codigo:"ANG36-2.5x2.5x1/4",  descripcion:"Ángulo A36 2 1/2\" × 2 1/2\" × 1/4\"",  peso_unit: 36.609, unidad:"barra"},
  {codigo:"ANG36-2.5x2.5x5/16", descripcion:"Ángulo A36 2 1/2\" × 2 1/2\" × 5/16\"", peso_unit: 44.645, unidad:"barra"},
  {codigo:"ANG36-2.5x2.5x3/8",  descripcion:"Ángulo A36 2 1/2\" × 2 1/2\" × 3/8\"",  peso_unit: 52.681, unidad:"barra"},
  {codigo:"ANG36-3x3x1/4",   descripcion:"Ángulo A36 3\" × 3\" × 1/4\"",   peso_unit:  43.752, unidad:"barra"},
  {codigo:"ANG36-3x3x5/16",  descripcion:"Ángulo A36 3\" × 3\" × 5/16\"",  peso_unit:  54.467, unidad:"barra"},
  {codigo:"ANG36-3x3x3/8",   descripcion:"Ángulo A36 3\" × 3\" × 3/8\"",   peso_unit:  64.289, unidad:"barra"},
  {codigo:"ANG36-3x3x1/2",   descripcion:"Ángulo A36 3\" × 3\" × 1/2\"",   peso_unit:  83.932, unidad:"barra"},
  {codigo:"ANG36-4x4x1/4",   descripcion:"Ángulo A36 4\" × 4\" × 1/4\"",   peso_unit:  58.920, unidad:"barra"},
  {codigo:"ANG36-4x4x5/16",  descripcion:"Ángulo A36 4\" × 4\" × 5/16\"",  peso_unit:  73.551, unidad:"barra"},
  {codigo:"ANG36-4x4x3/8",   descripcion:"Ángulo A36 4\" × 4\" × 3/8\"",   peso_unit:  87.492, unidad:"barra"},
  {codigo:"ANG36-4x4x1/2",   descripcion:"Ángulo A36 4\" × 4\" × 1/2\"",   peso_unit: 114.270, unidad:"barra"},
  {codigo:"ANG36-4x4x5/8",   descripcion:"Ángulo A36 4\" × 4\" × 5/8\"",   peso_unit: 141.074, unidad:"barra"},
  {codigo:"ANG36-5x5x5/16",  descripcion:"Ángulo A36 5\" × 5\" × 5/16\"",  peso_unit:  92.693, unidad:"barra"},
  {codigo:"ANG36-5x5x3/8",   descripcion:"Ángulo A36 5\" × 5\" × 3/8\"",   peso_unit: 109.806, unidad:"barra"},
  {codigo:"ANG36-5x5x1/2",   descripcion:"Ángulo A36 5\" × 5\" × 1/2\"",   peso_unit: 144.916, unidad:"barra"},
  {codigo:"ANG36-5x5x5/8",   descripcion:"Ángulo A36 5\" × 5\" × 5/8\"",   peso_unit: 179.358, unidad:"barra"},
  {codigo:"ANG36-6x6x5/16",  descripcion:"Ángulo A36 6\" × 6\" × 5/16\"",  peso_unit: 111.834, unidad:"barra"},
  {codigo:"ANG36-6x6x3/8",   descripcion:"Ángulo A36 6\" × 6\" × 3/8\"",   peso_unit: 133.287, unidad:"barra"},
  {codigo:"ANG36-6x6x1/2",   descripcion:"Ángulo A36 6\" × 6\" × 1/2\"",   peso_unit: 176.041, unidad:"barra"},
  {codigo:"ANG36-6x6x5/8",   descripcion:"Ángulo A36 6\" × 6\" × 5/8\"",   peso_unit: 216.480, unidad:"barra"},
  {codigo:"ANG36-8x8x1/2",   descripcion:"Ángulo A36 8\" × 8\" × 1/2\"",   peso_unit: 236.160, unidad:"barra"},
  {codigo:"ANG36-8x8x5/8",   descripcion:"Ángulo A36 8\" × 8\" × 5/8\"",   peso_unit: 292.516, unidad:"barra"},
  {codigo:"ANG36-8x8x3/4",   descripcion:"Ángulo A36 8\" × 8\" × 3/4\"",   peso_unit: 347.552, unidad:"barra"},
],

"PLATINAS ASTM A-36 / ISO 1035-4": [
  {codigo:"PLAT-1/8x1/2",    descripcion:"Platina A36 1/8\" × 1/2\"",    peso_unit:   1.98, unidad:"barra"},
  {codigo:"PLAT-1/8x5/8",    descripcion:"Platina A36 1/8\" × 5/8\"",    peso_unit:   2.34, unidad:"barra"},
  {codigo:"PLAT-1/8x3/4",    descripcion:"Platina A36 1/8\" × 3/4\"",    peso_unit:   2.88, unidad:"barra"},
  {codigo:"PLAT-1/8x1",      descripcion:"Platina A36 1/8\" × 1\"",      peso_unit:   3.84, unidad:"barra"},
  {codigo:"PLAT-1/8x1-1/4",  descripcion:"Platina A36 1/8\" × 1 1/4\"",  peso_unit:   4.80, unidad:"barra"},
  {codigo:"PLAT-1/8x1-1/2",  descripcion:"Platina A36 1/8\" × 1 1/2\"",  peso_unit:   5.82, unidad:"barra"},
  {codigo:"PLAT-1/8x2",      descripcion:"Platina A36 1/8\" × 2\"",      peso_unit:   7.68, unidad:"barra"},
  {codigo:"PLAT-3/16x1/2",   descripcion:"Platina A36 3/16\" × 1/2\"",   peso_unit:   2.88, unidad:"barra"},
  {codigo:"PLAT-3/16x5/8",   descripcion:"Platina A36 3/16\" × 5/8\"",   peso_unit:   3.66, unidad:"barra"},
  {codigo:"PLAT-3/16x3/4",   descripcion:"Platina A36 3/16\" × 3/4\"",   peso_unit:   4.44, unidad:"barra"},
  {codigo:"PLAT-3/16x1",     descripcion:"Platina A36 3/16\" × 1\"",     peso_unit:   5.88, unidad:"barra"},
  {codigo:"PLAT-3/16x1-1/4", descripcion:"Platina A36 3/16\" × 1 1/4\"", peso_unit:   7.38, unidad:"barra"},
  {codigo:"PLAT-3/16x1-1/2", descripcion:"Platina A36 3/16\" × 1 1/2\"", peso_unit:   9.36, unidad:"barra"},
  {codigo:"PLAT-3/16x2",     descripcion:"Platina A36 3/16\" × 2\"",     peso_unit:  11.40, unidad:"barra"},
  {codigo:"PLAT-3/16x2-1/2", descripcion:"Platina A36 3/16\" × 2 1/2\"", peso_unit:  14.28, unidad:"barra"},
  {codigo:"PLAT-1/4x1/2",    descripcion:"Platina A36 1/4\" × 1/2\"",    peso_unit:   3.84, unidad:"barra"},
  {codigo:"PLAT-1/4x5/8",    descripcion:"Platina A36 1/4\" × 5/8\"",    peso_unit:   4.80, unidad:"barra"},
  {codigo:"PLAT-1/4x3/4",    descripcion:"Platina A36 1/4\" × 3/4\"",    peso_unit:   5.70, unidad:"barra"},
  {codigo:"PLAT-1/4x1",      descripcion:"Platina A36 1/4\" × 1\"",      peso_unit:   7.68, unidad:"barra"},
  {codigo:"PLAT-1/4x1-1/4",  descripcion:"Platina A36 1/4\" × 1 1/4\"",  peso_unit:   9.66, unidad:"barra"},
  {codigo:"PLAT-1/4x1-1/2",  descripcion:"Platina A36 1/4\" × 1 1/2\"",  peso_unit:  11.52, unidad:"barra"},
  {codigo:"PLAT-1/4x2",      descripcion:"Platina A36 1/4\" × 2\"",      peso_unit:  15.24, unidad:"barra"},
  {codigo:"PLAT-1/4x2-1/2",  descripcion:"Platina A36 1/4\" × 2 1/2\"",  peso_unit:  19.08, unidad:"barra"},
  {codigo:"PLAT-1/4x3",      descripcion:"Platina A36 1/4\" × 3\"",      peso_unit:  22.86, unidad:"barra"},
  {codigo:"PLAT-1/4x4",      descripcion:"Platina A36 1/4\" × 4\"",      peso_unit:  30.36, unidad:"barra"},
  {codigo:"PLAT-3/8x1",      descripcion:"Platina A36 3/8\" × 1\"",      peso_unit:  11.52, unidad:"barra"},
  {codigo:"PLAT-3/8x1-1/4",  descripcion:"Platina A36 3/8\" × 1 1/4\"",  peso_unit:  14.28, unidad:"barra"},
  {codigo:"PLAT-3/8x1-1/2",  descripcion:"Platina A36 3/8\" × 1 1/2\"",  peso_unit:  17.22, unidad:"barra"},
  {codigo:"PLAT-3/8x2",      descripcion:"Platina A36 3/8\" × 2\"",      peso_unit:  22.92, unidad:"barra"},
  {codigo:"PLAT-3/8x2-1/2",  descripcion:"Platina A36 3/8\" × 2 1/2\"",  peso_unit:  28.62, unidad:"barra"},
  {codigo:"PLAT-3/8x3",      descripcion:"Platina A36 3/8\" × 3\"",      peso_unit:  34.44, unidad:"barra"},
  {codigo:"PLAT-3/8x4",      descripcion:"Platina A36 3/8\" × 4\"",      peso_unit:  45.60, unidad:"barra"},
  {codigo:"PLAT-1/2x1",      descripcion:"Platina A36 1/2\" × 1\"",      peso_unit:  15.24, unidad:"barra"},
  {codigo:"PLAT-1/2x1-1/2",  descripcion:"Platina A36 1/2\" × 1 1/2\"",  peso_unit:  23.04, unidad:"barra"},
  {codigo:"PLAT-1/2x2",      descripcion:"Platina A36 1/2\" × 2\"",      peso_unit:  30.54, unidad:"barra"},
  {codigo:"PLAT-1/2x2-1/2",  descripcion:"Platina A36 1/2\" × 2 1/2\"",  peso_unit:  38.10, unidad:"barra"},
  {codigo:"PLAT-1/2x3",      descripcion:"Platina A36 1/2\" × 3\"",      peso_unit:  45.78, unidad:"barra"},
  {codigo:"PLAT-1/2x4",      descripcion:"Platina A36 1/2\" × 4\"",      peso_unit:  61.02, unidad:"barra"},
  {codigo:"PLAT-1/2x5",      descripcion:"Platina A36 1/2\" × 5\"",      peso_unit:  76.567,unidad:"barra"},
  {codigo:"PLAT-1/2x6",      descripcion:"Platina A36 1/2\" × 6\"",      peso_unit:  91.879,unidad:"barra"},
  {codigo:"PLAT-5/8x3",      descripcion:"Platina A36 5/8\" × 3\"",      peso_unit:  58.229,unidad:"barra"},
  {codigo:"PLAT-5/8x4",      descripcion:"Platina A36 5/8\" × 4\"",      peso_unit:  77.639,unidad:"barra"},
  {codigo:"PLAT-5/8x5",      descripcion:"Platina A36 5/8\" × 5\"",      peso_unit:  97.048,unidad:"barra"},
  {codigo:"PLAT-5/8x6",      descripcion:"Platina A36 5/8\" × 6\"",      peso_unit: 116.458,unidad:"barra"},
  {codigo:"PLAT-3/4x3",      descripcion:"Platina A36 3/4\" × 3\"",      peso_unit:  68.371,unidad:"barra"},
  {codigo:"PLAT-3/4x4",      descripcion:"Platina A36 3/4\" × 4\"",      peso_unit:  91.161,unidad:"barra"},
  {codigo:"PLAT-3/4x5",      descripcion:"Platina A36 3/4\" × 5\"",      peso_unit: 113.951,unidad:"barra"},
  {codigo:"PLAT-3/4x6",      descripcion:"Platina A36 3/4\" × 6\"",      peso_unit: 136.742,unidad:"barra"},
  {codigo:"PLAT-1x3",        descripcion:"Platina A36 1\" × 3\"",        peso_unit:  91.161,unidad:"barra"},
  {codigo:"PLAT-1x4",        descripcion:"Platina A36 1\" × 4\"",        peso_unit: 121.548,unidad:"barra"},
  {codigo:"PLAT-1x5",        descripcion:"Platina A36 1\" × 5\"",        peso_unit: 151.935,unidad:"barra"},
  {codigo:"PLAT-1x6",        descripcion:"Platina A36 1\" × 6\"",        peso_unit: 182.322,unidad:"barra"},
],

"TEES ASTM A-36": [
  {codigo:"TEE-1/8x3/4x3/4",    descripcion:"Tee A36 1/8\" × 3/4\" × 3/4\"",       peso_unit:  5.28, unidad:"barra"},
  {codigo:"TEE-1/8x1x1",        descripcion:"Tee A36 1/8\" × 1\" × 1\"",            peso_unit:  6.60, unidad:"barra"},
  {codigo:"TEE-1/8x1-1/4x1-1/4",descripcion:"Tee A36 1/8\" × 1 1/4\" × 1 1/4\"",  peso_unit:  9.24, unidad:"barra"},
  {codigo:"TEE-1/8x1-1/2x1-1/2",descripcion:"Tee A36 1/8\" × 1 1/2\" × 1 1/2\"",  peso_unit: 11.04, unidad:"barra"},
  {codigo:"TEE-3/16x1-1/2x1-1/2",descripcion:"Tee A36 3/16\" × 1 1/2\" × 1 1/2\"",peso_unit: 16.32, unidad:"barra"},
  {codigo:"TEE-1/4x2x2",        descripcion:"Tee A36 1/4\" × 2\" × 2\"",           peso_unit: 29.82, unidad:"barra"},
],

"BARRAS CUADRADAS ASTM A-36": [
  {codigo:"BCUAD-1/4",   descripcion:"Barra Cuadrada A36 1/4\" × 1/4\"",  peso_unit:  1.899, unidad:"barra"},
  {codigo:"BCUAD-3/4",   descripcion:"Barra Cuadrada A36 3/4\" × 3/4\"",  peso_unit: 17.093, unidad:"barra"},
  {codigo:"BCUAD-7/8",   descripcion:"Barra Cuadrada A36 7/8\" × 7/8\"",  peso_unit: 23.265, unidad:"barra"},
  {codigo:"BCUAD-1",     descripcion:"Barra Cuadrada A36 1\" × 1\"",      peso_unit: 30.387, unidad:"barra"},
  {codigo:"BCUAD-9mm",   descripcion:"Barra Cuadrada A36 9 × 9 mm",       peso_unit:  3.840, unidad:"mm"},
  {codigo:"BCUAD-12mm",  descripcion:"Barra Cuadrada A36 12 × 12 mm",     peso_unit:  6.780, unidad:"mm"},
  {codigo:"BCUAD-15mm",  descripcion:"Barra Cuadrada A36 15 × 15 mm",     peso_unit: 10.620, unidad:"mm"},
],

"BARRAS CUADRADAS ORNAMENTALES ASTM A-36": [
  {codigo:"BORN-9mm",    descripcion:"Barra Cuadrada Ornamental A36 9 mm",  peso_unit:  3.26, unidad:"mm"},
  {codigo:"BORN-12mm",   descripcion:"Barra Cuadrada Ornamental A36 12 mm", peso_unit:  5.52, unidad:"mm"},
  {codigo:"BORN-15mm",   descripcion:"Barra Cuadrada Ornamental A36 15 mm", peso_unit:  8.59, unidad:"mm"},
],

"BARRAS REDONDAS LISAS ASTM A-36": [
  {codigo:"BRED-3/8",    descripcion:"Barra Redonda Lisa A36 3/8\" × 20 pies",   peso_unit:   3.36, unidad:"pies"},
  {codigo:"BRED-1/2",    descripcion:"Barra Redonda Lisa A36 1/2\" × 20 pies",   peso_unit:   5.97, unidad:"pies"},
  {codigo:"BRED-5/8",    descripcion:"Barra Redonda Lisa A36 5/8\" × 20 pies",   peso_unit:   9.33, unidad:"pies"},
  {codigo:"BRED-3/4",    descripcion:"Barra Redonda Lisa A36 3/4\" × 20 pies",   peso_unit:  13.44, unidad:"pies"},
  {codigo:"BRED-7/8",    descripcion:"Barra Redonda Lisa A36 7/8\" × 20 pies",   peso_unit:  18.28, unidad:"pies"},
  {codigo:"BRED-1",      descripcion:"Barra Redonda Lisa A36 1\" × 20 pies",     peso_unit:  23.87, unidad:"pies"},
  {codigo:"BRED-1-1/8",  descripcion:"Barra Redonda Lisa A36 1 1/8\" × 20 pies", peso_unit:  30.21, unidad:"pies"},
  {codigo:"BRED-1-1/4",  descripcion:"Barra Redonda Lisa A36 1 1/4\" × 20 pies", peso_unit:  37.30, unidad:"pies"},
  {codigo:"BRED-1-3/8",  descripcion:"Barra Redonda Lisa A36 1 3/8\" × 20 pies", peso_unit:  45.13, unidad:"pies"},
  {codigo:"BRED-1-1/2",  descripcion:"Barra Redonda Lisa A36 1 1/2\" × 20 pies", peso_unit:  53.70, unidad:"pies"},
  {codigo:"BRED-1-3/4",  descripcion:"Barra Redonda Lisa A36 1 3/4\" × 20 pies", peso_unit:  73.09, unidad:"pies"},
  {codigo:"BRED-2",      descripcion:"Barra Redonda Lisa A36 2\" × 20 pies",     peso_unit:  95.47, unidad:"pies"},
  {codigo:"BRED-2-1/4",  descripcion:"Barra Redonda Lisa A36 2 1/4\" × 20 pies", peso_unit: 120.83, unidad:"pies"},
  {codigo:"BRED-2-1/2",  descripcion:"Barra Redonda Lisa A36 2 1/2\" × 20 pies", peso_unit: 149.16, unidad:"pies"},
],

"BARRAS CALIBRADAS REDONDAS (SAE 1022 / SAE 1213)": [
  {codigo:"BCAL-RED-1/4",    descripcion:"Barra Calibrada Redonda 1/4\" (Cal.1016)",      peso_unit:  1.492, unidad:"barra"},
  {codigo:"BCAL-RED-5/16",   descripcion:"Barra Calibrada Redonda 5/16\" (Cal.1213/1022)",peso_unit:  2.331, unidad:"barra"},
  {codigo:"BCAL-RED-3/8",    descripcion:"Barra Calibrada Redonda 3/8\" (Cal.1213/1022)", peso_unit:  3.356, unidad:"barra"},
  {codigo:"BCAL-RED-7/16",   descripcion:"Barra Calibrada Redonda 7/16\" (Cal.1213/1022)",peso_unit:  4.568, unidad:"barra"},
  {codigo:"BCAL-RED-1/2",    descripcion:"Barra Calibrada Redonda 1/2\" (Cal.1213/1022)", peso_unit:  5.966, unidad:"barra"},
  {codigo:"BCAL-RED-9/16",   descripcion:"Barra Calibrada Redonda 9/16\" (Cal.1213/1022)",peso_unit:  7.551, unidad:"barra"},
  {codigo:"BCAL-RED-5/8",    descripcion:"Barra Calibrada Redonda 5/8\" (Cal.1213/1022)", peso_unit:  9.323, unidad:"barra"},
  {codigo:"BCAL-RED-3/4",    descripcion:"Barra Calibrada Redonda 3/4\" (Cal.1213/1022)", peso_unit: 13.425, unidad:"barra"},
  {codigo:"BCAL-RED-7/8",    descripcion:"Barra Calibrada Redonda 7/8\" (Cal.1213/1022)", peso_unit: 18.272, unidad:"barra"},
  {codigo:"BCAL-RED-1",      descripcion:"Barra Calibrada Redonda 1\" (Cal.1213/1022)",   peso_unit: 23.866, unidad:"barra"},
  {codigo:"BCAL-RED-1-1/8",  descripcion:"Barra Calibrada Redonda 1 1/8\" (Cal.1022)",    peso_unit: 30.205, unidad:"barra"},
  {codigo:"BCAL-RED-1-1/4",  descripcion:"Barra Calibrada Redonda 1 1/4\" (Cal.1022)",    peso_unit: 37.291, unidad:"barra"},
  {codigo:"BCAL-RED-1-3/8",  descripcion:"Barra Calibrada Redonda 1 3/8\" (Cal.1022)",    peso_unit: 45.122, unidad:"barra"},
  {codigo:"BCAL-RED-1-1/2",  descripcion:"Barra Calibrada Redonda 1 1/2\" (Cal.1022)",    peso_unit: 53.698, unidad:"barra"},
  {codigo:"BCAL-RED-1-3/4",  descripcion:"Barra Calibrada Redonda 1 3/4\" (Cal.1022)",    peso_unit: 73.090, unidad:"barra"},
  {codigo:"BCAL-RED-2",      descripcion:"Barra Calibrada Redonda 2\" (Cal.1022)",         peso_unit: 95.464, unidad:"barra"},
  {codigo:"BCAL-RED-2-1/4",  descripcion:"Barra Calibrada Redonda 2 1/4\" (Cal.1022)",    peso_unit:120.822, unidad:"barra"},
  {codigo:"BCAL-RED-2-1/2",  descripcion:"Barra Calibrada Redonda 2 1/2\" (Cal.1022)",    peso_unit:149.162, unidad:"barra"},
],

"BARRAS CALIBRADAS HEXAGONALES (SAE 1022 / SAE 1213)": [
  {codigo:"BCAL-HEX-11/16",  descripcion:"Barra Calibrada Hexagonal 11/16\" (Cal.1022)",      peso_unit: 12.438, unidad:"barra"},
  {codigo:"BCAL-HEX-3/4",    descripcion:"Barra Calibrada Hexagonal 3/4\" (Cal.1213/1022)",   peso_unit: 14.802, unidad:"barra"},
  {codigo:"BCAL-HEX-13/16",  descripcion:"Barra Calibrada Hexagonal 13/16\" (Cal.1213/1022)", peso_unit: 17.372, unidad:"barra"},
  {codigo:"BCAL-HEX-7/8",    descripcion:"Barra Calibrada Hexagonal 7/8\" (Cal.1213/1022)",   peso_unit: 20.148, unidad:"barra"},
  {codigo:"BCAL-HEX-15/16",  descripcion:"Barra Calibrada Hexagonal 15/16\" (Cal.1213/1022)", peso_unit: 23.129, unidad:"barra"},
  {codigo:"BCAL-HEX-1",      descripcion:"Barra Calibrada Hexagonal 1\" (Cal.1213)",          peso_unit: 26.315, unidad:"barra"},
  {codigo:"BCAL-HEX-1-1/16", descripcion:"Barra Calibrada Hexagonal 1 1/16\" (Cal.1213/1022)",peso_unit: 29.707, unidad:"barra"},
  {codigo:"BCAL-HEX-1-1/8",  descripcion:"Barra Calibrada Hexagonal 1 1/8\" (Cal.1213)",      peso_unit: 33.305, unidad:"barra"},
  {codigo:"BCAL-HEX-1-1/4",  descripcion:"Barra Calibrada Hexagonal 1 1/4\" (Cal.1213)",      peso_unit: 41.117, unidad:"barra"},
  {codigo:"BCAL-HEX-1-3/8",  descripcion:"Barra Calibrada Hexagonal 1 3/8\" (Cal.1213/1022)", peso_unit: 49.752, unidad:"barra"},
  {codigo:"BCAL-HEX-1-1/2",  descripcion:"Barra Calibrada Hexagonal 1 1/2\" (Cal.1213)",      peso_unit: 59.209, unidad:"barra"},
],

"BARRAS HEXAGONALES SAE 1045": [
  {codigo:"BHEX1045-1",      descripcion:"Barra Hexagonal SAE 1045 1\"",      peso_unit: 26.315, unidad:"barra"},
  {codigo:"BHEX1045-1-1/8",  descripcion:"Barra Hexagonal SAE 1045 1 1/8\"",  peso_unit: 33.305, unidad:"barra"},
  {codigo:"BHEX1045-1-1/4",  descripcion:"Barra Hexagonal SAE 1045 1 1/4\"",  peso_unit: 41.117, unidad:"barra"},
],

"CANAL U ASTM A-36": [
  {codigo:"CANAL-2x1",        descripcion:"Canal U A36 2\" × 1\" ala",          peso_unit:  2.58, unidad:"barra"},
  {codigo:"CANAL-3x1-3/8",    descripcion:"Canal U A36 3\" × 1 3/8\" ala",      peso_unit:  4.10, unidad:"barra"},
  {codigo:"CANAL-3x1-1/2",    descripcion:"Canal U A36 3\" × 1 1/2\" ala",      peso_unit:  5.00, unidad:"barra"},
  {codigo:"CANAL-3x1-5/8",    descripcion:"Canal U A36 3\" × 1 5/8\" ala",      peso_unit:  6.00, unidad:"barra"},
  {codigo:"CANAL-4x1-5/8",    descripcion:"Canal U A36 4\" × 1 5/8\" ala",      peso_unit:  5.40, unidad:"barra"},
  {codigo:"CANAL-4x1-3/4",    descripcion:"Canal U A36 4\" × 1 3/4\" ala",      peso_unit:  7.25, unidad:"barra"},
  {codigo:"CANAL-5x1-3/4",    descripcion:"Canal U A36 5\" × 1 3/4\" ala",      peso_unit:  6.70, unidad:"barra"},
  {codigo:"CANAL-5x1-7/8",    descripcion:"Canal U A36 5\" × 1 7/8\" ala",      peso_unit:  9.00, unidad:"barra"},
  {codigo:"CANAL-6x1-7/8",    descripcion:"Canal U A36 6\" × 1 7/8\" ala",      peso_unit:  8.20, unidad:"barra"},
  {codigo:"CANAL-6x2",        descripcion:"Canal U A36 6\" × 2\" ala",           peso_unit: 10.50, unidad:"barra"},
  {codigo:"CANAL-6x2-1/8",    descripcion:"Canal U A36 6\" × 2 1/8\" ala",      peso_unit: 13.00, unidad:"barra"},
  {codigo:"CANAL-7x2-1/8",    descripcion:"Canal U A36 7\" × 2 1/8\" ala",      peso_unit:  9.80, unidad:"barra"},
  {codigo:"CANAL-7x2-1/4a",   descripcion:"Canal U A36 7\" × 2 1/4\" ala (12.25 lb/pie)", peso_unit: 12.25, unidad:"barra"},
  {codigo:"CANAL-7x2-1/4b",   descripcion:"Canal U A36 7\" × 2 1/4\" ala (14.75 lb/pie)", peso_unit: 14.75, unidad:"barra"},
  {codigo:"CANAL-8x1-7/8",    descripcion:"Canal U A36 8\" × 1 7/8\" ala",      peso_unit:  8.50, unidad:"barra"},
  {codigo:"CANAL-8x2-1/4",    descripcion:"Canal U A36 8\" × 2 1/4\" ala",      peso_unit: 11.50, unidad:"barra"},
  {codigo:"CANAL-8x2-3/8",    descripcion:"Canal U A36 8\" × 2 3/8\" ala",      peso_unit: 13.75, unidad:"barra"},
  {codigo:"CANAL-8x2-1/2",    descripcion:"Canal U A36 8\" × 2 1/2\" ala",      peso_unit: 18.75, unidad:"barra"},
  {codigo:"CANAL-9x2-3/8",    descripcion:"Canal U A36 9\" × 2 3/8\" ala",      peso_unit: 13.40, unidad:"barra"},
  {codigo:"CANAL-9x2-1/2",    descripcion:"Canal U A36 9\" × 2 1/2\" ala",      peso_unit: 15.00, unidad:"barra"},
  {codigo:"CANAL-9x2-5/8",    descripcion:"Canal U A36 9\" × 2 5/8\" ala",      peso_unit: 20.00, unidad:"barra"},
  {codigo:"CANAL-10x2-5/8",   descripcion:"Canal U A36 10\" × 2 5/8\" ala",     peso_unit: 15.30, unidad:"barra"},
  {codigo:"CANAL-10x2-3/4",   descripcion:"Canal U A36 10\" × 2 3/4\" ala",     peso_unit: 20.00, unidad:"barra"},
  {codigo:"CANAL-10x2-7/8",   descripcion:"Canal U A36 10\" × 2 7/8\" ala",     peso_unit: 25.00, unidad:"barra"},
  {codigo:"CANAL-10x3",       descripcion:"Canal U A36 10\" × 3\" ala",          peso_unit: 30.00, unidad:"barra"},
  {codigo:"CANAL-12x3a",      descripcion:"Canal U A36 12\" × 3\" ala (20.7 lb/pie)",  peso_unit: 20.70, unidad:"barra"},
  {codigo:"CANAL-12x3b",      descripcion:"Canal U A36 12\" × 3\" ala (25.0 lb/pie)",  peso_unit: 25.00, unidad:"barra"},
  {codigo:"CANAL-12x3-1/8",   descripcion:"Canal U A36 12\" × 3 1/8\" ala",     peso_unit: 30.00, unidad:"barra"},
  {codigo:"CANAL-15x3-3/8",   descripcion:"Canal U A36 15\" × 3 3/8\" ala",     peso_unit: 33.90, unidad:"barra"},
  {codigo:"CANAL-15x3-1/2",   descripcion:"Canal U A36 15\" × 3 1/2\" ala",     peso_unit: 40.00, unidad:"barra"},
  {codigo:"CANAL-15x3-3/4",   descripcion:"Canal U A36 15\" × 3 3/4\" ala",     peso_unit: 50.00, unidad:"barra"},
  {codigo:"CANAL-18x4a",      descripcion:"Canal U A36 18\" × 4\" ala (42.7 lb/pie)",  peso_unit: 42.70, unidad:"barra"},
  {codigo:"CANAL-18x4b",      descripcion:"Canal U A36 18\" × 4\" ala (45.8 lb/pie)",  peso_unit: 45.80, unidad:"barra"},
  {codigo:"CANAL-18x4-1/8",   descripcion:"Canal U A36 18\" × 4 1/8\" ala",     peso_unit: 51.90, unidad:"barra"},
  {codigo:"CANAL-18x4-1/4",   descripcion:"Canal U A36 18\" × 4 1/4\" ala",     peso_unit: 58.00, unidad:"barra"},
],
  
"VIGA I STANDARD AMERICANO ASTM A-36": [
    { codigo:"S3x5.7",    descripcion:'Viga I A36 3" × 2 3/8" ala – 5.70 lbs/pie',   ala_b:"2 3/8", peso_unit:  5.70, alma_w:  4.318, ala_f:  6.604, unidad:"barra" },
    { codigo:"S3x7.5",    descripcion:'Viga I A36 3" × 2 1/2" ala – 7.50 lbs/pie',   ala_b:"2 1/2", peso_unit:  7.50, alma_w:  8.854, ala_f:  6.604, unidad:"barra" },
    { codigo:"S4x7.7",    descripcion:'Viga I A36 4" × 2 5/8" ala – 7.70 lbs/pie',   ala_b:"2 5/8", peso_unit:  7.70, alma_w:  4.826, ala_f:  7.442, unidad:"barra" },
    { codigo:"S4x9.5",    descripcion:'Viga I A36 4" × 2 3/4" ala – 9.50 lbs/pie',   ala_b:"2 3/4", peso_unit:  9.50, alma_w:  8.280, ala_f:  7.442, unidad:"barra" },
    { codigo:"S5x10",     descripcion:'Viga I A36 5" × 3" ala – 10.00 lbs/pie',       ala_b:"3",     peso_unit: 10.00, alma_w:  5.334, ala_f:  8.280, unidad:"barra" },
    { codigo:"S5x14.75",  descripcion:'Viga I A36 5" × 3 1/4" ala – 14.75 lbs/pie',  ala_b:"3 1/4", peso_unit: 14.75, alma_w: 12.547, ala_f:  8.280, unidad:"barra" },
    { codigo:"S6x12.5",   descripcion:'Viga I A36 6" × 3 3/8" ala – 12.50 lbs/pie',  ala_b:"3 3/8", peso_unit: 12.50, alma_w:  5.842, ala_f:  9.118, unidad:"barra" },
    { codigo:"S6x17.25",  descripcion:'Viga I A36 6" × 3 5/8" ala – 17.25 lbs/pie',  ala_b:"3 5/8", peso_unit: 17.25, alma_w: 11.811, ala_f:  9.118, unidad:"barra" },
    { codigo:"S7x15.3",   descripcion:'Viga I A36 7" × 3 5/8" ala – 15.30 lbs/pie',  ala_b:"3 5/8", peso_unit: 15.30, alma_w:  6.350, ala_f:  9.956, unidad:"barra" },
    { codigo:"S7x20",     descripcion:'Viga I A36 7" × 3 7/8" ala – 20.00 lbs/pie',  ala_b:"3 7/8", peso_unit: 20.00, alma_w: 11.430, ala_f:  9.956, unidad:"barra" },
    { codigo:"S8x18.4",   descripcion:'Viga I A36 8" × 4" ala – 18.40 lbs/pie',      ala_b:"4",     peso_unit: 18.40, alma_w:  6.858, ala_f: 10.795, unidad:"barra" },
    { codigo:"S8x23",     descripcion:'Viga I A36 8" × 4 1/8" ala – 23.00 lbs/pie',  ala_b:"4 1/8", peso_unit: 23.00, alma_w: 11.201, ala_f: 10.795, unidad:"barra" },
    { codigo:"S10x25.4",  descripcion:'Viga I A36 10" × 4 5/8" ala – 25.40 lbs/pie', ala_b:"4 5/8", peso_unit: 25.40, alma_w:  7.874, ala_f: 12.471, unidad:"barra" },
    { codigo:"S10x35",    descripcion:'Viga I A36 10" × 5" ala – 35.00 lbs/pie',     ala_b:"5",     peso_unit: 35.00, alma_w: 15.087, ala_f: 12.471, unidad:"barra" },
    { codigo:"S12x31.8",  descripcion:'Viga I A36 12" × 5" ala – 31.80 lbs/pie',     ala_b:"5",     peso_unit: 31.80, alma_w:  8.890, ala_f: 13.817, unidad:"barra" },
    { codigo:"S12x35",    descripcion:'Viga I A36 12" × 5 1/8" ala – 35.00 lbs/pie', ala_b:"5 1/8", peso_unit: 35.00, alma_w: 10.871, ala_f: 13.817, unidad:"barra" },
    { codigo:"S12x40.8",  descripcion:'Viga I A36 12" × 5 1/4" ala – 40.80 lbs/pie', ala_b:"5 1/4", peso_unit: 40.80, alma_w: 11.684, ala_f: 16.738, unidad:"barra" },
    { codigo:"S12x50",    descripcion:'Viga I A36 12" × 5 1/2" ala – 50.00 lbs/pie', ala_b:"5 1/2", peso_unit: 50.00, alma_w: 17.499, ala_f: 16.738, unidad:"barra" },
    { codigo:"S15x42.9",  descripcion:'Viga I A36 15" × 5 1/2" ala – 42.90 lbs/pie', ala_b:"5 1/2", peso_unit: 42.90, alma_w: 10.414, ala_f: 15.798, unidad:"barra" },
    { codigo:"S15x50",    descripcion:'Viga I A36 15" × 5 5/8" ala – 50.00 lbs/pie', ala_b:"5 5/8", peso_unit: 50.00, alma_w: 13.970, ala_f: 15.798, unidad:"barra" },
    { codigo:"S18x54.7",  descripcion:'Viga I A36 18" × 6" ala – 54.70 lbs/pie',     ala_b:"6",     peso_unit: 54.70, alma_w: 11.684, ala_f: 17.551, unidad:"barra" },
    { codigo:"S18x70",    descripcion:'Viga I A36 18" × 6 1/4" ala – 70.00 lbs/pie', ala_b:"6 1/4", peso_unit: 70.00, alma_w: 18.059, ala_f: 17.551, unidad:"barra" },
    { codigo:"S20x65.4",  descripcion:'Viga I A36 20" × 6 1/4" ala – 65.40 lbs/pie', ala_b:"6 1/4", peso_unit: 65.40, alma_w: 12.700, ala_f: 20.040, unidad:"barra" },
    { codigo:"S20x75",    descripcion:'Viga I A36 20" × 6 3/8" ala – 75.00 lbs/pie', ala_b:"6 3/8", peso_unit: 75.00, alma_w: 16.281, ala_f: 20.040, unidad:"barra" },
    { codigo:"S20x85",    descripcion:'Viga I A36 20" × 7" ala – 85.00 lbs/pie',     ala_b:"7",     peso_unit: 85.00, alma_w: 16.586, ala_f: 23.266, unidad:"barra" },
    { codigo:"S20x95",    descripcion:'Viga I A36 20" × 7 1/4" ala – 95.00 lbs/pie', ala_b:"7 1/4", peso_unit: 95.00, alma_w: 20.320, ala_f: 23.266, unidad:"barra" },
    { codigo:"S24x79.9",  descripcion:'Viga I A36 24" × 7" ala – 79.90 lbs/pie',     ala_b:"7",     peso_unit:  79.90, alma_w: 12.700, ala_f: 22.123, unidad:"barra" },
    { codigo:"S24x90",    descripcion:'Viga I A36 24" × 7 1/8" ala – 90.00 lbs/pie', ala_b:"7 1/8", peso_unit:  90.00, alma_w: 15.849, ala_f: 22.123, unidad:"barra" },
    { codigo:"S24x100",   descripcion:'Viga I A36 24" × 7 1/4" ala – 100.0 lbs/pie', ala_b:"7 1/4", peso_unit: 100.00, alma_w: 18.973, ala_f: 19.837, unidad:"barra" },
    { codigo:"S24x105.9", descripcion:'Viga I A36 24" × 7 7/8" ala – 105.9 lbs/pie', ala_b:"7 7/8", peso_unit: 105.90, alma_w: 15.875, ala_f: 27.990, unidad:"barra" },
    { codigo:"S24x120",   descripcion:'Viga I A36 24" × 8" ala – 120.0 lbs/pie',     ala_b:"8",     peso_unit: 120.00, alma_w: 20.269, ala_f: 27.990, unidad:"barra" },
  ],

"VIGA H WIDE FLANGE BEAMS ASTM A36 A572 GRADO 50": [
    { codigo:"W4x13",     descripcion:'Viga H A36/A572-50 4" × 4" ala – 13.0 lbs/pie',      ala_b:"4",      peso_unit:  13.0, alma_w:  7.112, ala_f:  8.763, unidad:"barra" },
    { codigo:"W5x16",     descripcion:'Viga H A36/A572-50 5" × 5" ala – 16.0 lbs/pie',      ala_b:"5",      peso_unit:  16.0, alma_w:  6.096, ala_f:  9.144, unidad:"barra" },
    { codigo:"W5x19",     descripcion:'Viga H A36/A572-50 5" × 5" ala – 19.0 lbs/pie',      ala_b:"5",      peso_unit:  19.0, alma_w:  6.858, ala_f: 10.922, unidad:"barra" },
    { codigo:"W6x9",      descripcion:'Viga H A36/A572-50 6" × 4" ala – 9.0 lbs/pie',       ala_b:"4",      peso_unit:   9.0, alma_w:  4.318, ala_f:  5.461, unidad:"barra" },
    { codigo:"W6x12",     descripcion:'Viga H A36/A572-50 6" × 4" ala – 12.0 lbs/pie',      ala_b:"4",      peso_unit:  12.0, alma_w:  5.842, ala_f:  7.112, unidad:"barra" },
    { codigo:"W6x16",     descripcion:'Viga H A36/A572-50 6" × 4" ala – 16.0 lbs/pie',      ala_b:"4",      peso_unit:  16.0, alma_w:  6.604, ala_f: 10.287, unidad:"barra" },
    { codigo:"W6x15",     descripcion:'Viga H A36/A572-50 6" × 6" ala – 15.0 lbs/pie',      ala_b:"6",      peso_unit:  15.0, alma_w:  5.842, ala_f:  6.604, unidad:"barra" },
    { codigo:"W6x20",     descripcion:'Viga H A36/A572-50 6" × 6" ala – 20.0 lbs/pie',      ala_b:"6",      peso_unit:  20.0, alma_w:  6.604, ala_f:  9.271, unidad:"barra" },
    { codigo:"W6x25",     descripcion:'Viga H A36/A572-50 6" × 6 1/8" ala – 25.0 lbs/pie',  ala_b:"6 1/8",  peso_unit:  25.0, alma_w:  8.128, ala_f: 11.557, unidad:"barra" },
    { codigo:"W8x10",     descripcion:'Viga H A36/A572-50 8" × 4" ala – 10.0 lbs/pie',      ala_b:"4",      peso_unit:  10.0, alma_w:  4.318, ala_f:  5.207, unidad:"barra" },
    { codigo:"W8x13",     descripcion:'Viga H A36/A572-50 8" × 4" ala – 13.0 lbs/pie',      ala_b:"4",      peso_unit:  13.0, alma_w:  5.842, ala_f:  6.477, unidad:"barra" },
    { codigo:"W8x15",     descripcion:'Viga H A36/A572-50 8" × 4" ala – 15.0 lbs/pie',      ala_b:"4",      peso_unit:  15.0, alma_w:  6.223, ala_f:  8.001, unidad:"barra" },
    { codigo:"W8x18",     descripcion:'Viga H A36/A572-50 8" × 5 1/4" ala – 18.0 lbs/pie',  ala_b:"5 1/4",  peso_unit:  18.0, alma_w:  5.842, ala_f:  8.382, unidad:"barra" },
    { codigo:"W8x21",     descripcion:'Viga H A36/A572-50 8" × 5 1/4" ala – 21.0 lbs/pie',  ala_b:"5 1/4",  peso_unit:  21.0, alma_w:  6.350, ala_f: 10.160, unidad:"barra" },
    { codigo:"W8x24",     descripcion:'Viga H A36/A572-50 8" × 6 1/2" ala – 24.0 lbs/pie',  ala_b:"6 1/2",  peso_unit:  24.0, alma_w:  6.223, ala_f: 10.160, unidad:"barra" },
    { codigo:"W8x28",     descripcion:'Viga H A36/A572-50 8" × 6 1/2" ala – 28.0 lbs/pie',  ala_b:"6 1/2",  peso_unit:  28.0, alma_w:  7.239, ala_f: 11.811, unidad:"barra" },
    { codigo:"W8x31",     descripcion:'Viga H A36/A572-50 8" × 8" ala – 31.0 lbs/pie',      ala_b:"8",      peso_unit:  31.0, alma_w:  7.239, ala_f: 11.049, unidad:"barra" },
    { codigo:"W8x35",     descripcion:'Viga H A36/A572-50 8" × 8" ala – 35.0 lbs/pie',      ala_b:"8",      peso_unit:  35.0, alma_w:  7.874, ala_f: 12.573, unidad:"barra" },
    { codigo:"W8x40",     descripcion:'Viga H A36/A572-50 8" × 8 1/8" ala – 40.0 lbs/pie',  ala_b:"8 1/8",  peso_unit:  40.0, alma_w:  9.144, ala_f: 14.224, unidad:"barra" },
    { codigo:"W8x48",     descripcion:'Viga H A36/A572-50 8" × 8 1/8" ala – 48.0 lbs/pie',  ala_b:"8 1/8",  peso_unit:  48.0, alma_w: 10.160, ala_f: 17.399, unidad:"barra" },
    { codigo:"W8x58",     descripcion:'Viga H A36/A572-50 8" × 8 1/4" ala – 58.0 lbs/pie',  ala_b:"8 1/4",  peso_unit:  58.0, alma_w: 12.954, ala_f: 20.574, unidad:"barra" },
    { codigo:"W8x67",     descripcion:'Viga H A36/A572-50 8" × 8 1/4" ala – 67.0 lbs/pie',  ala_b:"8 1/4",  peso_unit:  67.0, alma_w: 14.478, ala_f: 23.749, unidad:"barra" },
    { codigo:"W10x12",    descripcion:'Viga H A36/A572-50 10" × 4" ala – 12.0 lbs/pie',     ala_b:"4",      peso_unit:  12.0, alma_w:  4.826, ala_f:  5.334, unidad:"barra" },
    { codigo:"W10x15",    descripcion:'Viga H A36/A572-50 10" × 4" ala – 15.0 lbs/pie',     ala_b:"4",      peso_unit:  15.0, alma_w:  5.842, ala_f:  6.858, unidad:"barra" },
    { codigo:"W10x17",    descripcion:'Viga H A36/A572-50 10" × 4" ala – 17.0 lbs/pie',     ala_b:"4",      peso_unit:  17.0, alma_w:  6.096, ala_f:  8.382, unidad:"barra" },
    { codigo:"W10x19",    descripcion:'Viga H A36/A572-50 10" × 4" ala – 19.0 lbs/pie',     ala_b:"4",      peso_unit:  19.0, alma_w:  6.350, ala_f: 10.033, unidad:"barra" },
    { codigo:"W10x22",    descripcion:'Viga H A36/A572-50 10" × 5 3/4" ala – 22.0 lbs/pie', ala_b:"5 3/4",  peso_unit:  22.0, alma_w:  6.096, ala_f:  9.144, unidad:"barra" },
    { codigo:"W10x26",    descripcion:'Viga H A36/A572-50 10" × 5 3/4" ala – 26.0 lbs/pie', ala_b:"5 3/4",  peso_unit:  26.0, alma_w:  6.604, ala_f: 11.176, unidad:"barra" },
    { codigo:"W10x30",    descripcion:'Viga H A36/A572-50 10" × 5 3/4" ala – 30.0 lbs/pie', ala_b:"5 3/4",  peso_unit:  30.0, alma_w:  7.620, ala_f: 12.954, unidad:"barra" },
    { codigo:"W10x33",    descripcion:'Viga H A36/A572-50 10" × 8" ala – 33.0 lbs/pie',     ala_b:"8",      peso_unit:  33.0, alma_w:  7.366, ala_f: 11.049, unidad:"barra" },
    { codigo:"W10x39",    descripcion:'Viga H A36/A572-50 10" × 8" ala – 39.0 lbs/pie',     ala_b:"8",      peso_unit:  39.0, alma_w:  8.001, ala_f: 13.462, unidad:"barra" },
    { codigo:"W10x45",    descripcion:'Viga H A36/A572-50 10" × 8" ala – 45.0 lbs/pie',     ala_b:"8",      peso_unit:  45.0, alma_w:  8.890, ala_f: 15.748, unidad:"barra" },
    { codigo:"W10x49",    descripcion:'Viga H A36/A572-50 10" × 10" ala – 49.0 lbs/pie',    ala_b:"10",     peso_unit:  49.0, alma_w:  8.636, ala_f: 14.224, unidad:"barra" },
    { codigo:"W10x54",    descripcion:'Viga H A36/A572-50 10" × 10" ala – 54.0 lbs/pie',    ala_b:"10",     peso_unit:  54.0, alma_w:  9.398, ala_f: 15.621, unidad:"barra" },
    { codigo:"W10x60",    descripcion:'Viga H A36/A572-50 10" × 10 1/8" ala – 60.0 lbs/pie',ala_b:"10 1/8", peso_unit:  60.0, alma_w: 10.668, ala_f: 17.272, unidad:"barra" },
    { codigo:"W10x68",    descripcion:'Viga H A36/A572-50 10" × 10 1/8" ala – 68.0 lbs/pie',ala_b:"10 1/8", peso_unit:  68.0, alma_w: 11.938, ala_f: 19.558, unidad:"barra" },
    { codigo:"W10x77",    descripcion:'Viga H A36/A572-50 10" × 10 1/4" ala – 77.0 lbs/pie',ala_b:"10 1/4", peso_unit:  77.0, alma_w: 13.462, ala_f: 22.098, unidad:"barra" },
    { codigo:"W10x88",    descripcion:'Viga H A36/A572-50 10" × 10 1/4" ala – 88.0 lbs/pie',ala_b:"10 1/4", peso_unit:  88.0, alma_w: 15.367, ala_f: 25.146, unidad:"barra" },
    { codigo:"W10x100",   descripcion:'Viga H A36/A572-50 10" × 10 3/8" ala – 100.0 lbs/pie',ala_b:"10 3/8",peso_unit: 100.0, alma_w: 17.399, ala_f: 28.448, unidad:"barra" },
    { codigo:"W10x112",   descripcion:'Viga H A36/A572-50 10" × 10 3/8" ala – 112.0 lbs/pie',ala_b:"10 3/8",peso_unit: 112.0, alma_w: 19.177, ala_f: 31.750, unidad:"barra" },
    { codigo:"W12x14",    descripcion:'Viga H A36/A572-50 12" × 4" ala – 14.0 lbs/pie',     ala_b:"4",      peso_unit:  14.0, alma_w:  5.080, ala_f:  5.715, unidad:"barra" },
    { codigo:"W12x16",    descripcion:'Viga H A36/A572-50 12" × 4" ala – 16.0 lbs/pie',     ala_b:"4",      peso_unit:  16.0, alma_w:  5.588, ala_f:  6.731, unidad:"barra" },
    { codigo:"W12x19",    descripcion:'Viga H A36/A572-50 12" × 4" ala – 19.0 lbs/pie',     ala_b:"4",      peso_unit:  19.0, alma_w:  5.969, ala_f:  8.890, unidad:"barra" },
    { codigo:"W12x22",    descripcion:'Viga H A36/A572-50 12" × 4" ala – 22.0 lbs/pie',     ala_b:"4",      peso_unit:  22.0, alma_w:  6.604, ala_f: 10.795, unidad:"barra" },
    { codigo:"W12x26",    descripcion:'Viga H A36/A572-50 12" × 6 1/2" ala – 26.0 lbs/pie', ala_b:"6 1/2",  peso_unit:  26.0, alma_w:  5.842, ala_f:  9.652, unidad:"barra" },
    { codigo:"W12x30",    descripcion:'Viga H A36/A572-50 12" × 6 1/2" ala – 30.0 lbs/pie', ala_b:"6 1/2",  peso_unit:  30.0, alma_w:  6.604, ala_f: 11.176, unidad:"barra" },
    { codigo:"W12x35",    descripcion:'Viga H A36/A572-50 12" × 6 1/2" ala – 35.0 lbs/pie', ala_b:"6 1/2",  peso_unit:  35.0, alma_w:  7.620, ala_f: 13.208, unidad:"barra" },
    { codigo:"W12x40",    descripcion:'Viga H A36/A572-50 12" × 8" ala – 40.0 lbs/pie',     ala_b:"8",      peso_unit:  40.0, alma_w:  7.493, ala_f: 13.081, unidad:"barra" },
    { codigo:"W12x45",    descripcion:'Viga H A36/A572-50 12" × 8" ala – 45.0 lbs/pie',     ala_b:"8",      peso_unit:  45.0, alma_w:  8.509, ala_f: 14.605, unidad:"barra" },
    { codigo:"W12x50",    descripcion:'Viga H A36/A572-50 12" × 8 1/8" ala – 50.0 lbs/pie', ala_b:"8 1/8",  peso_unit:  50.0, alma_w:  9.398, ala_f: 16.256, unidad:"barra" },
    { codigo:"W12x53",    descripcion:'Viga H A36/A572-50 12" × 10" ala – 53.0 lbs/pie',    ala_b:"10",     peso_unit:  53.0, alma_w:  8.763, ala_f: 14.605, unidad:"barra" },
    { codigo:"W12x58",    descripcion:'Viga H A36/A572-50 12" × 10" ala – 58.0 lbs/pie',    ala_b:"10",     peso_unit:  58.0, alma_w:  9.144, ala_f: 16.256, unidad:"barra" },
    { codigo:"W12x65",    descripcion:'Viga H A36/A572-50 12" × 12" ala – 65.0 lbs/pie',    ala_b:"12",     peso_unit:  65.0, alma_w:  9.906, ala_f: 15.367, unidad:"barra" },
    { codigo:"W12x72",    descripcion:'Viga H A36/A572-50 12" × 12" ala – 72.0 lbs/pie',    ala_b:"12",     peso_unit:  72.0, alma_w: 10.922, ala_f: 17.018, unidad:"barra" },
    { codigo:"W12x79",    descripcion:'Viga H A36/A572-50 12" × 12 1/8" ala – 79.0 lbs/pie', ala_b:"12 1/8", peso_unit:  79.0, alma_w: 11.938, ala_f: 18.669, unidad:"barra" },
    { codigo:"W12x87",    descripcion:'Viga H A36/A572-50 12" × 12 1/8" ala – 87.0 lbs/pie', ala_b:"12 1/8", peso_unit:  87.0, alma_w: 13.081, ala_f: 20.574, unidad:"barra" },
    { codigo:"W12x96",    descripcion:'Viga H A36/A572-50 12" × 12 1/8" ala – 96.0 lbs/pie', ala_b:"12 1/8", peso_unit:  96.0, alma_w: 13.970, ala_f: 22.860, unidad:"barra" },
    { codigo:"W12x106",   descripcion:'Viga H A36/A572-50 12" × 12 1/4" ala – 106.0 lbs/pie',ala_b:"12 1/4", peso_unit: 106.0, alma_w: 15.494, ala_f: 25.146, unidad:"barra" },
    { codigo:"W12x120",   descripcion:'Viga H A36/A572-50 12" × 12 3/8" ala – 120.0 lbs/pie',ala_b:"12 3/8", peso_unit: 120.0, alma_w: 18.034, ala_f: 28.067, unidad:"barra" },
    { codigo:"W12x136",   descripcion:'Viga H A36/A572-50 12" × 12 3/8" ala – 136.0 lbs/pie',ala_b:"12 3/8", peso_unit: 136.0, alma_w: 20.066, ala_f: 31.750, unidad:"barra" },
    { codigo:"W12x152",   descripcion:'Viga H A36/A572-50 12" × 12 1/2" ala – 152.0 lbs/pie',ala_b:"12 1/2", peso_unit: 152.0, alma_w: 22.098, ala_f: 35.560, unidad:"barra" },
    { codigo:"W12x170",   descripcion:'Viga H A36/A572-50 12" × 12 5/8" ala – 170.0 lbs/pie',ala_b:"12 5/8", peso_unit: 170.0, alma_w: 24.384, ala_f: 39.624, unidad:"barra" },
    { codigo:"W12x190",   descripcion:'Viga H A36/A572-50 12" × 12 5/8" ala – 190.0 lbs/pie',ala_b:"12 5/8", peso_unit: 190.0, alma_w: 26.924, ala_f: 44.069, unidad:"barra" },
    { codigo:"W14x22",    descripcion:'Viga H A36/A572-50 14" × 5" ala – 22.0 lbs/pie',     ala_b:"5",      peso_unit:  22.0, alma_w:  5.842, ala_f:  8.509, unidad:"barra" },
    { codigo:"W14x26",    descripcion:'Viga H A36/A572-50 14" × 5" ala – 26.0 lbs/pie',     ala_b:"5",      peso_unit:  26.0, alma_w:  6.477, ala_f: 10.668, unidad:"barra" },
    { codigo:"W14x30",    descripcion:'Viga H A36/A572-50 14" × 6 3/4" ala – 30.0 lbs/pie', ala_b:"6 3/4",  peso_unit:  30.0, alma_w:  6.858, ala_f:  9.779, unidad:"barra" },
    { codigo:"W14x34",    descripcion:'Viga H A36/A572-50 14" × 6 3/4" ala – 34.0 lbs/pie', ala_b:"6 3/4",  peso_unit:  34.0, alma_w:  7.239, ala_f: 11.557, unidad:"barra" },
    { codigo:"W14x38",    descripcion:'Viga H A36/A572-50 14" × 6 3/4" ala – 38.0 lbs/pie', ala_b:"6 3/4",  peso_unit:  38.0, alma_w:  7.874, ala_f: 13.081, unidad:"barra" },
    { codigo:"W14x43",    descripcion:'Viga H A36/A572-50 14" × 8" ala – 43.0 lbs/pie',     ala_b:"8",      peso_unit:  43.0, alma_w:  7.747, ala_f: 13.462, unidad:"barra" },
    { codigo:"W14x48",    descripcion:'Viga H A36/A572-50 14" × 8" ala – 48.0 lbs/pie',     ala_b:"8",      peso_unit:  48.0, alma_w:  8.636, ala_f: 15.113, unidad:"barra" },
    { codigo:"W14x53",    descripcion:'Viga H A36/A572-50 14" × 8" ala – 53.0 lbs/pie',     ala_b:"8",      peso_unit:  53.0, alma_w:  9.398, ala_f: 16.764, unidad:"barra" },
    { codigo:"W14x61",    descripcion:'Viga H A36/A572-50 14" × 10" ala – 61.0 lbs/pie',    ala_b:"10",     peso_unit:  61.0, alma_w:  9.525, ala_f: 16.383, unidad:"barra" },
    { codigo:"W14x68",    descripcion:'Viga H A36/A572-50 14" × 10" ala – 68.0 lbs/pie',    ala_b:"10",     peso_unit:  68.0, alma_w: 10.541, ala_f: 18.288, unidad:"barra" },
    { codigo:"W14x74",    descripcion:'Viga H A36/A572-50 14" × 10 1/8" ala – 74.0 lbs/pie', ala_b:"10 1/8", peso_unit:  74.0, alma_w: 11.430, ala_f: 19.939, unidad:"barra" },
    { codigo:"W14x82",    descripcion:'Viga H A36/A572-50 14" × 10 1/8" ala – 82.0 lbs/pie', ala_b:"10 1/8", peso_unit:  82.0, alma_w: 12.954, ala_f: 21.717, unidad:"barra" },
    { codigo:"W14x90",    descripcion:'Viga H A36/A572-50 14" × 14 1/2" ala – 90.0 lbs/pie', ala_b:"14 1/2", peso_unit:  90.0, alma_w: 11.176, ala_f: 18.034, unidad:"barra" },
    { codigo:"W14x99",    descripcion:'Viga H A36/A572-50 14" × 14 5/8" ala – 99.0 lbs/pie', ala_b:"14 5/8", peso_unit:  99.0, alma_w: 12.319, ala_f: 19.812, unidad:"barra" },
    { codigo:"W14x109",   descripcion:'Viga H A36/A572-50 14" × 14 5/8" ala – 109.0 lbs/pie',ala_b:"14 5/8", peso_unit: 109.0, alma_w: 13.335, ala_f: 21.844, unidad:"barra" },
    { codigo:"W14x120",   descripcion:'Viga H A36/A572-50 14" × 14 5/8" ala – 120.0 lbs/pie',ala_b:"14 5/8", peso_unit: 120.0, alma_w: 14.986, ala_f: 23.876, unidad:"barra" },
    { codigo:"W14x132",   descripcion:'Viga H A36/A572-50 14" × 14 3/4" ala – 132.0 lbs/pie',ala_b:"14 3/4", peso_unit: 132.0, alma_w: 16.332, ala_f: 26.162, unidad:"barra" },
    { codigo:"W14x145",   descripcion:'Viga H A36/A572-50 14" × 15 1/2" ala – 145.0 lbs/pie',ala_b:"15 1/2", peso_unit: 145.0, alma_w: 17.272, ala_f: 27.686, unidad:"barra" },
    { codigo:"W14x159",   descripcion:'Viga H A36/A572-50 14" × 15 5/8" ala – 159.0 lbs/pie',ala_b:"15 5/8", peso_unit: 159.0, alma_w: 18.923, ala_f: 30.226, unidad:"barra" },
    { codigo:"W14x176",   descripcion:'Viga H A36/A572-50 14" × 15 5/8" ala – 176.0 lbs/pie',ala_b:"15 5/8", peso_unit: 176.0, alma_w: 21.082, ala_f: 33.274, unidad:"barra" },
    { codigo:"W14x193",   descripcion:'Viga H A36/A572-50 14" × 15 3/4" ala – 193.0 lbs/pie',ala_b:"15 3/4", peso_unit: 193.0, alma_w: 22.606, ala_f: 36.576, unidad:"barra" },
    { codigo:"W14x211",   descripcion:'Viga H A36/A572-50 14" × 15 3/4" ala – 211.0 lbs/pie',ala_b:"15 3/4", peso_unit: 211.0, alma_w: 24.892, ala_f: 39.624, unidad:"barra" },
    { codigo:"W14x233",   descripcion:'Viga H A36/A572-50 14" × 15 7/8" ala – 233.0 lbs/pie',ala_b:"15 7/8", peso_unit: 233.0, alma_w: 27.178, ala_f: 43.688, unidad:"barra" },
    { codigo:"W14x257",   descripcion:'Viga H A36/A572-50 14" × 16" ala – 257.0 lbs/pie',   ala_b:"16",     peso_unit: 257.0, alma_w: 29.845, ala_f: 48.006, unidad:"barra" },
    { codigo:"W14x283",   descripcion:'Viga H A36/A572-50 14" × 16 1/8" ala – 283.0 lbs/pie',ala_b:"16 1/8", peso_unit: 283.0, alma_w: 32.766, ala_f: 52.578, unidad:"barra" },
    { codigo:"W14x311",   descripcion:'Viga H A36/A572-50 14" × 16 1/4" ala – 311.0 lbs/pie',ala_b:"16 1/4", peso_unit: 311.0, alma_w: 35.814, ala_f: 57.404, unidad:"barra" },
    { codigo:"W14x342",   descripcion:'Viga H A36/A572-50 14" × 16 3/8" ala – 342.0 lbs/pie',ala_b:"16 3/8", peso_unit: 342.0, alma_w: 39.116, ala_f: 62.738, unidad:"barra" },
    { codigo:"W14x370",   descripcion:'Viga H A36/A572-50 14" × 16 1/2" ala – 370.0 lbs/pie',ala_b:"16 1/2", peso_unit: 370.0, alma_w: 42.037, ala_f: 67.564, unidad:"barra" },
    { codigo:"W14x398",   descripcion:'Viga H A36/A572-50 14" × 16 5/8" ala – 398.0 lbs/pie',ala_b:"16 5/8", peso_unit: 398.0, alma_w: 44.958, ala_f: 72.263, unidad:"barra" },
    { codigo:"W14x426",   descripcion:'Viga H A36/A572-50 14" × 16 3/4" ala – 426.0 lbs/pie',ala_b:"16 3/4", peso_unit: 426.0, alma_w: 47.625, ala_f: 77.089, unidad:"barra" },
    { codigo:"W14x455",   descripcion:'Viga H A36/A572-50 14" × 16 7/8" ala – 455.0 lbs/pie',ala_b:"16 7/8", peso_unit: 455.0, alma_w: 51.181, ala_f: 81.534, unidad:"barra" },
    { codigo:"W14x500",   descripcion:'Viga H A36/A572-50 14" × 17" ala – 500.0 lbs/pie',   ala_b:"17",     peso_unit: 500.0, alma_w: 55.626, ala_f: 88.900, unidad:"barra" },
    { codigo:"W14x550",   descripcion:'Viga H A36/A572-50 14" × 17 1/4" ala – 550.0 lbs/pie',ala_b:"17 1/4", peso_unit: 550.0, alma_w: 60.452, ala_f: 97.028, unidad:"barra" },
    { codigo:"W14x605",   descripcion:'Viga H A36/A572-50 14" × 17 3/8" ala – 605.0 lbs/pie',ala_b:"17 3/8", peso_unit: 605.0, alma_w: 65.913, ala_f:105.664, unidad:"barra" },
    { codigo:"W14x665",   descripcion:'Viga H A36/A572-50 14" × 17 5/8" ala – 665.0 lbs/pie',ala_b:"17 5/8", peso_unit: 665.0, alma_w: 71.882, ala_f:114.808, unidad:"barra" },
    { codigo:"W14x730",   descripcion:'Viga H A36/A572-50 14" × 17 7/8" ala – 730.0 lbs/pie',ala_b:"17 7/8", peso_unit: 730.0, alma_w: 77.978, ala_f:124.714, unidad:"barra" },
    { codigo:"W16x26",    descripcion:'Viga H A36/A572-50 16" × 5 1/2" ala – 26.0 lbs/pie', ala_b:"5 1/2",  peso_unit:  26.0, alma_w:  5.350, ala_f:  8.763, unidad:"barra" },
    { codigo:"W16x31",    descripcion:'Viga H A36/A572-50 16" × 5 1/2" ala – 31.0 lbs/pie', ala_b:"5 1/2",  peso_unit:  31.0, alma_w:  5.985, ala_f: 11.176, unidad:"barra" },
    { codigo:"W16x36",    descripcion:'Viga H A36/A572-50 16" × 7" ala – 36.0 lbs/pie',     ala_b:"7",      peso_unit:  36.0, alma_w:  7.493, ala_f: 10.922, unidad:"barra" },
    { codigo:"W16x40",    descripcion:'Viga H A36/A572-50 16" × 7" ala – 40.0 lbs/pie',     ala_b:"7",      peso_unit:  40.0, alma_w:  7.747, ala_f: 12.827, unidad:"barra" },
    { codigo:"W16x45",    descripcion:'Viga H A36/A572-50 16" × 7" ala – 45.0 lbs/pie',     ala_b:"7",      peso_unit:  45.0, alma_w:  8.763, ala_f: 14.351, unidad:"barra" },
    { codigo:"W16x50",    descripcion:'Viga H A36/A572-50 16" × 7 1/8" ala – 50.0 lbs/pie', ala_b:"7 1/8",  peso_unit:  50.0, alma_w:  9.652, ala_f: 16.002, unidad:"barra" },
    { codigo:"W16x57",    descripcion:'Viga H A36/A572-50 16" × 7 1/8" ala – 57.0 lbs/pie', ala_b:"7 1/8",  peso_unit:  57.0, alma_w: 10.922, ala_f: 18.161, unidad:"barra" },
    { codigo:"W16x67",    descripcion:'Viga H A36/A572-50 16" × 10 1/4" ala – 67.0 lbs/pie', ala_b:"10 1/4", peso_unit:  67.0, alma_w: 10.033, ala_f: 16.891, unidad:"barra" },
    { codigo:"W16x77",    descripcion:'Viga H A36/A572-50 16" × 10 1/4" ala – 77.0 lbs/pie', ala_b:"10 1/4", peso_unit:  77.0, alma_w: 11.557, ala_f: 19.304, unidad:"barra" },
    { codigo:"W16x89",    descripcion:'Viga H A36/A572-50 16" × 10 3/8" ala – 89.0 lbs/pie', ala_b:"10 3/8", peso_unit:  89.0, alma_w: 13.335, ala_f: 22.225, unidad:"barra" },
    { codigo:"W16x100",   descripcion:'Viga H A36/A572-50 16" × 10 3/8" ala – 100.0 lbs/pie',ala_b:"10 3/8", peso_unit: 100.0, alma_w: 14.859, ala_f: 25.019, unidad:"barra" },
    { codigo:"W18x35",    descripcion:'Viga H A36/A572-50 18" × 6" ala – 35.0 lbs/pie',     ala_b:"6",      peso_unit:  35.0, alma_w:  7.620, ala_f: 10.795, unidad:"barra" },
    { codigo:"W18x40",    descripcion:'Viga H A36/A572-50 18" × 6" ala – 40.0 lbs/pie',     ala_b:"6",      peso_unit:  40.0, alma_w:  8.001, ala_f: 13.335, unidad:"barra" },
    { codigo:"W18x45",    descripcion:'Viga H A36/A572-50 18" × 6" ala – 45.0 lbs/pie',     ala_b:"6",      peso_unit:  45.0, alma_w:  9.144, ala_f: 15.367, unidad:"barra" },
    { codigo:"W18x50",    descripcion:'Viga H A36/A572-50 18" × 7 1/2" ala – 50.0 lbs/pie', ala_b:"7 1/2",  peso_unit:  50.0, alma_w:  9.017, ala_f: 14.478, unidad:"barra" },
    { codigo:"W18x55",    descripcion:'Viga H A36/A572-50 18" × 7 1/2" ala – 55.0 lbs/pie', ala_b:"7 1/2",  peso_unit:  55.0, alma_w:  9.906, ala_f: 16.002, unidad:"barra" },
    { codigo:"W18x60",    descripcion:'Viga H A36/A572-50 18" × 7 1/2" ala – 60.0 lbs/pie', ala_b:"7 1/2",  peso_unit:  60.0, alma_w: 10.541, ala_f: 17.653, unidad:"barra" },
    { codigo:"W18x65",    descripcion:'Viga H A36/A572-50 18" × 7 5/8" ala – 65.0 lbs/pie', ala_b:"7 5/8",  peso_unit:  65.0, alma_w: 11.430, ala_f: 19.050, unidad:"barra" },
    { codigo:"W18x71",    descripcion:'Viga H A36/A572-50 18" × 7 5/8" ala – 71.0 lbs/pie', ala_b:"7 5/8",  peso_unit:  71.0, alma_w: 12.573, ala_f: 20.574, unidad:"barra" },
    { codigo:"W18x76",    descripcion:'Viga H A36/A572-50 18" × 11" ala – 76.0 lbs/pie',    ala_b:"11",     peso_unit:  76.0, alma_w: 10.795, ala_f: 17.272, unidad:"barra" },
    { codigo:"W18x86",    descripcion:'Viga H A36/A572-50 18" × 11 1/2" ala – 86.0 lbs/pie', ala_b:"11 1/2", peso_unit:  86.0, alma_w: 12.192, ala_f: 19.558, unidad:"barra" },
    { codigo:"W18x97",    descripcion:'Viga H A36/A572-50 18" × 11 1/2" ala – 97.0 lbs/pie', ala_b:"11 1/2", peso_unit:  97.0, alma_w: 13.589, ala_f: 22.098, unidad:"barra" },
    { codigo:"W18x106",   descripcion:'Viga H A36/A572-50 18" × 11 1/4" ala – 106.0 lbs/pie',ala_b:"11 1/4", peso_unit: 106.0, alma_w: 14.986, ala_f: 23.876, unidad:"barra" },
    { codigo:"W18x119",   descripcion:'Viga H A36/A572-50 18" × 11 1/4" ala – 119.0 lbs/pie',ala_b:"11 1/4", peso_unit: 119.0, alma_w: 16.637, ala_f: 26.924, unidad:"barra" },
    { codigo:"W21x44",    descripcion:'Viga H A36/A572-50 21" × 6 1/2" ala – 44.0 lbs/pie', ala_b:"6 1/2",  peso_unit:  44.0, alma_w:  8.890, ala_f: 11.430, unidad:"barra" },
    { codigo:"W21x50",    descripcion:'Viga H A36/A572-50 21" × 6 1/2" ala – 50.0 lbs/pie', ala_b:"6 1/2",  peso_unit:  50.0, alma_w:  9.652, ala_f: 13.589, unidad:"barra" },
    { codigo:"W21x57",    descripcion:'Viga H A36/A572-50 21" × 6 1/2" ala – 57.0 lbs/pie', ala_b:"6 1/2",  peso_unit:  57.0, alma_w: 10.287, ala_f: 16.510, unidad:"barra" },
    { codigo:"W21x62",    descripcion:'Viga H A36/A572-50 21" × 8 1/4" ala – 62.0 lbs/pie', ala_b:"8 1/4",  peso_unit:  62.0, alma_w: 10.160, ala_f: 15.621, unidad:"barra" },
    { codigo:"W21x68",    descripcion:'Viga H A36/A572-50 21" × 8 1/4" ala – 68.0 lbs/pie', ala_b:"8 1/4",  peso_unit:  68.0, alma_w: 10.922, ala_f: 17.399, unidad:"barra" },
    { codigo:"W21x73",    descripcion:'Viga H A36/A572-50 21" × 8 1/4" ala – 73.0 lbs/pie', ala_b:"8 1/4",  peso_unit:  73.0, alma_w: 11.557, ala_f: 18.796, unidad:"barra" },
    { codigo:"W21x83",    descripcion:'Viga H A36/A572-50 21" × 8 3/8" ala – 83.0 lbs/pie', ala_b:"8 3/8",  peso_unit:  83.0, alma_w: 13.081, ala_f: 21.209, unidad:"barra" },
    { codigo:"W21x93",    descripcion:'Viga H A36/A572-50 21" × 8 3/8" ala – 93.0 lbs/pie', ala_b:"8 3/8",  peso_unit:  93.0, alma_w: 14.732, ala_f: 23.622, unidad:"barra" },
    { codigo:"W21x101",   descripcion:'Viga H A36/A572-50 21" × 12 1/4" ala – 101.0 lbs/pie',ala_b:"12 1/4", peso_unit: 101.0, alma_w: 12.700, ala_f: 20.320, unidad:"barra" },
    { codigo:"W21x111",   descripcion:'Viga H A36/A572-50 21" × 12 3/8" ala – 111.0 lbs/pie',ala_b:"12 3/8", peso_unit: 111.0, alma_w: 14.097, ala_f: 22.225, unidad:"barra" },
    { codigo:"W21x122",   descripcion:'Viga H A36/A572-50 21" × 12 3/8" ala – 122.0 lbs/pie',ala_b:"12 3/8", peso_unit: 122.0, alma_w: 15.240, ala_f: 24.384, unidad:"barra" },
    { codigo:"W21x132",   descripcion:'Viga H A36/A572-50 21" × 12 1/2" ala – 132.0 lbs/pie',ala_b:"12 1/2", peso_unit: 132.0, alma_w: 16.510, ala_f: 26.289, unidad:"barra" },
    { codigo:"W21x147",   descripcion:'Viga H A36/A572-50 21" × 12 1/2" ala – 147.0 lbs/pie',ala_b:"12 1/2", peso_unit: 147.0, alma_w: 18.288, ala_f: 29.210, unidad:"barra" },
    { codigo:"W24x55",    descripcion:'Viga H A36/A572-50 24" × 7" ala – 55.0 lbs/pie',     ala_b:"7",      peso_unit:  55.0, alma_w: 10.033, ala_f: 12.827, unidad:"barra" },
    { codigo:"W24x62",    descripcion:'Viga H A36/A572-50 24" × 7" ala – 62.0 lbs/pie',     ala_b:"7",      peso_unit:  62.0, alma_w: 10.033, ala_f: 14.986, unidad:"barra" },
    { codigo:"W24x68",    descripcion:'Viga H A36/A572-50 24" × 9" ala – 68.0 lbs/pie',     ala_b:"9",      peso_unit:  68.0, alma_w: 10.541, ala_f: 14.859, unidad:"barra" },
    { codigo:"W24x76",    descripcion:'Viga H A36/A572-50 24" × 9" ala – 76.0 lbs/pie',     ala_b:"9",      peso_unit:  76.0, alma_w: 11.176, ala_f: 17.272, unidad:"barra" },
    { codigo:"W24x84",    descripcion:'Viga H A36/A572-50 24" × 9" ala – 84.0 lbs/pie',     ala_b:"9",      peso_unit:  84.0, alma_w: 11.938, ala_f: 19.558, unidad:"barra" },
    { codigo:"W24x94",    descripcion:'Viga H A36/A572-50 24" × 9 1/8" ala – 94.0 lbs/pie', ala_b:"9 1/8",  peso_unit:  94.0, alma_w: 13.081, ala_f: 22.225, unidad:"barra" },
    { codigo:"W24x104",   descripcion:'Viga H A36/A572-50 24" × 12 3/4" ala – 104.0 lbs/pie',ala_b:"12 3/4", peso_unit: 104.0, alma_w: 12.700, ala_f: 19.050, unidad:"barra" },
    { codigo:"W24x117",   descripcion:'Viga H A36/A572-50 24" × 12 3/4" ala – 117.0 lbs/pie',ala_b:"12 3/4", peso_unit: 117.0, alma_w: 13.970, ala_f: 21.590, unidad:"barra" },
    { codigo:"W24x131",   descripcion:'Viga H A36/A572-50 24" × 12 7/8" ala – 131.0 lbs/pie',ala_b:"12 7/8", peso_unit: 131.0, alma_w: 15.367, ala_f: 24.384, unidad:"barra" },
    { codigo:"W24x146",   descripcion:'Viga H A36/A572-50 24" × 12 7/8" ala – 146.0 lbs/pie',ala_b:"12 7/8", peso_unit: 146.0, alma_w: 16.510, ala_f: 27.686, unidad:"barra" },
    { codigo:"W24x162",   descripcion:'Viga H A36/A572-50 24" × 13" ala – 162.0 lbs/pie',   ala_b:"13",     peso_unit: 162.0, alma_w: 17.907, ala_f: 30.988, unidad:"barra" },
    { codigo:"W27x84",    descripcion:'Viga H A36/A572-50 27" × 10" ala – 84.0 lbs/pie',    ala_b:"10",     peso_unit:  84.0, alma_w: 11.684, ala_f: 16.256, unidad:"barra" },
    { codigo:"W27x94",    descripcion:'Viga H A36/A572-50 27" × 10" ala – 94.0 lbs/pie',    ala_b:"10",     peso_unit:  94.0, alma_w: 12.446, ala_f: 18.923, unidad:"barra" },
    { codigo:"W27x102",   descripcion:'Viga H A36/A572-50 27" × 10" ala – 102.0 lbs/pie',   ala_b:"10",     peso_unit: 102.0, alma_w: 13.081, ala_f: 21.082, unidad:"barra" },
    { codigo:"W27x114",   descripcion:'Viga H A36/A572-50 27" × 10 1/8" ala – 114.0 lbs/pie',ala_b:"10 1/8", peso_unit: 114.0, alma_w: 14.478, ala_f: 23.622, unidad:"barra" },
    { codigo:"W27x146",   descripcion:'Viga H A36/A572-50 27" × 14" ala – 146.0 lbs/pie',   ala_b:"14",     peso_unit: 146.0, alma_w: 15.367, ala_f: 24.765, unidad:"barra" },
    { codigo:"W27x161",   descripcion:'Viga H A36/A572-50 27" × 14" ala – 161.0 lbs/pie',   ala_b:"14",     peso_unit: 161.0, alma_w: 16.764, ala_f: 27.432, unidad:"barra" },
    { codigo:"W27x178",   descripcion:'Viga H A36/A572-50 27" × 14 1/8" ala – 178.0 lbs/pie',ala_b:"14 1/8", peso_unit: 178.0, alma_w: 18.415, ala_f: 30.226, unidad:"barra" },
    { codigo:"W30x99",    descripcion:'Viga H A36/A572-50 30" × 10 1/2" ala – 99.0 lbs/pie', ala_b:"10 1/2", peso_unit:  99.0, alma_w: 13.208, ala_f: 17.018, unidad:"barra" },
    { codigo:"W30x108",   descripcion:'Viga H A36/A572-50 30" × 10 1/2" ala – 108.0 lbs/pie',ala_b:"10 1/2", peso_unit: 108.0, alma_w: 13.843, ala_f: 19.304, unidad:"barra" },
    { codigo:"W30x116",   descripcion:'Viga H A36/A572-50 30" × 10 1/2" ala – 116.0 lbs/pie',ala_b:"10 1/2", peso_unit: 116.0, alma_w: 14.351, ala_f: 21.590, unidad:"barra" },
    { codigo:"W30x124",   descripcion:'Viga H A36/A572-50 30" × 10 1/2" ala – 124.0 lbs/pie',ala_b:"10 1/2", peso_unit: 124.0, alma_w: 14.859, ala_f: 23.622, unidad:"barra" },
    { codigo:"W30x132",   descripcion:'Viga H A36/A572-50 30" × 10 1/2" ala – 132.0 lbs/pie',ala_b:"10 1/2", peso_unit: 132.0, alma_w: 15.621, ala_f: 25.400, unidad:"barra" },
    { codigo:"W30x173",   descripcion:'Viga H A36/A572-50 30" × 15" ala – 173.0 lbs/pie',   ala_b:"15",     peso_unit: 173.0, alma_w: 16.637, ala_f: 27.051, unidad:"barra" },
    { codigo:"W30x191",   descripcion:'Viga H A36/A572-50 30" × 15" ala – 191.0 lbs/pie',   ala_b:"15",     peso_unit: 191.0, alma_w: 18.034, ala_f: 30.099, unidad:"barra" },
    { codigo:"W30x211",   descripcion:'Viga H A36/A572-50 30" × 15 1/8" ala – 211.0 lbs/pie',ala_b:"15 1/8", peso_unit: 211.0, alma_w: 19.685, ala_f: 33.401, unidad:"barra" },
    { codigo:"W33x118",   descripcion:'Viga H A36/A572-50 33" × 11 1/2" ala – 118.0 lbs/pie',ala_b:"11 1/2", peso_unit: 118.0, alma_w: 13.970, ala_f: 18.796, unidad:"barra" },
    { codigo:"W33x130",   descripcion:'Viga H A36/A572-50 33" × 11 1/2" ala – 130.0 lbs/pie',ala_b:"11 1/2", peso_unit: 130.0, alma_w: 14.732, ala_f: 21.717, unidad:"barra" },
    { codigo:"W33x141",   descripcion:'Viga H A36/A572-50 33" × 11 1/2" ala – 141.0 lbs/pie',ala_b:"11 1/2", peso_unit: 141.0, alma_w: 15.367, ala_f: 24.384, unidad:"barra" },
    { codigo:"W33x152",   descripcion:'Viga H A36/A572-50 33" × 11 5/8" ala – 152.0 lbs/pie',ala_b:"11 5/8", peso_unit: 152.0, alma_w: 16.129, ala_f: 26.797, unidad:"barra" },
    { codigo:"W33x201",   descripcion:'Viga H A36/A572-50 33" × 15 3/4" ala – 201.0 lbs/pie',ala_b:"15 3/4", peso_unit: 201.0, alma_w: 18.161, ala_f: 29.210, unidad:"barra" },
    { codigo:"W33x221",   descripcion:'Viga H A36/A572-50 33" × 15 3/4" ala – 221.0 lbs/pie',ala_b:"15 3/4", peso_unit: 221.0, alma_w: 19.685, ala_f: 32.385, unidad:"barra" },
    { codigo:"W33x241",   descripcion:'Viga H A36/A572-50 33" × 15 7/8" ala – 241.0 lbs/pie',ala_b:"15 7/8", peso_unit: 241.0, alma_w: 21.082, ala_f: 35.560, unidad:"barra" },
    { codigo:"W36x135",   descripcion:'Viga H A36/A572-50 36" × 12" ala – 135.0 lbs/pie',   ala_b:"12",     peso_unit: 135.0, alma_w: 15.240, ala_f: 20.066, unidad:"barra" },
    { codigo:"W36x150",   descripcion:'Viga H A36/A572-50 36" × 12" ala – 150.0 lbs/pie',   ala_b:"12",     peso_unit: 150.0, alma_w: 15.875, ala_f: 23.876, unidad:"barra" },
    { codigo:"W36x160",   descripcion:'Viga H A36/A572-50 36" × 12" ala – 160.0 lbs/pie',   ala_b:"12",     peso_unit: 160.0, alma_w: 16.510, ala_f: 25.908, unidad:"barra" },
    { codigo:"W36x170",   descripcion:'Viga H A36/A572-50 36" × 12" ala – 170.0 lbs/pie',   ala_b:"12",     peso_unit: 170.0, alma_w: 17.272, ala_f: 27.940, unidad:"barra" },
    { codigo:"W36x182",   descripcion:'Viga H A36/A572-50 36" × 12 1/8" ala – 182.0 lbs/pie',ala_b:"12 1/8", peso_unit: 182.0, alma_w: 18.415, ala_f: 29.972, unidad:"barra" },
    { codigo:"W36x194",   descripcion:'Viga H A36/A572-50 36" × 12 1/8" ala – 194.0 lbs/pie',ala_b:"12 1/8", peso_unit: 194.0, alma_w: 19.431, ala_f: 32.004, unidad:"barra" },
    { codigo:"W36x210",   descripcion:'Viga H A36/A572-50 36" × 12 1/8" ala – 210.0 lbs/pie',ala_b:"12 1/8", peso_unit: 210.0, alma_w: 21.082, ala_f: 34.544, unidad:"barra" },
    { codigo:"W36x230",   descripcion:'Viga H A36/A572-50 36" × 16 1/2" ala – 230.0 lbs/pie',ala_b:"16 1/2", peso_unit: 230.0, alma_w: 19.304, ala_f: 32.004, unidad:"barra" },
    { codigo:"W36x245",   descripcion:'Viga H A36/A572-50 36" × 16 1/2" ala – 245.0 lbs/pie',ala_b:"16 1/2", peso_unit: 245.0, alma_w: 20.320, ala_f: 34.290, unidad:"barra" },
    { codigo:"W36x260",   descripcion:'Viga H A36/A572-50 36" × 16 1/2" ala – 260.0 lbs/pie',ala_b:"16 1/2", peso_unit: 260.0, alma_w: 21.336, ala_f: 36.576, unidad:"barra" },
    { codigo:"W36x280",   descripcion:'Viga H A36/A572-50 36" × 16 5/8" ala – 280.0 lbs/pie',ala_b:"16 5/8", peso_unit: 280.0, alma_w: 22.479, ala_f: 39.878, unidad:"barra" },
    { codigo:"W36x300",   descripcion:'Viga H A36/A572-50 36" × 16 5/8" ala – 300.0 lbs/pie',ala_b:"16 5/8", peso_unit: 300.0, alma_w: 24.003, ala_f: 42.672, unidad:"barra" },
  ],

"BARRAS DE CONSTRUCCION ASTM A-615 GRADO 60 / ITINTEC 341.031 / ARN420-91": [
  {codigo:"BCON-1/4",    descripcion:"Barra Construcción G60 1/4\" (6.0mm)",   peso_unit: 0.222, unidad:"metro"},
  {codigo:"BCON-5/16",   descripcion:"Barra Construcción G60 5/16\" (8.0mm)",  peso_unit: 0.395, unidad:"metro"},
  {codigo:"BCON-3/8",    descripcion:"Barra Construcción G60 3/8\" (9.5mm)",   peso_unit: 0.560, unidad:"metro"},
  {codigo:"BCON-1/2-12", descripcion:"Barra Construcción G60 1/2\" (12.0mm)",  peso_unit: 0.888, unidad:"metro"},
  {codigo:"BCON-1/2-127",descripcion:"Barra Construcción G60 1/2\" (12.7mm)",  peso_unit: 0.994, unidad:"metro"},
  {codigo:"BCON-5/8",    descripcion:"Barra Construcción G60 5/8\" (15.9mm)",  peso_unit: 1.552, unidad:"metro"},
  {codigo:"BCON-3/4",    descripcion:"Barra Construcción G60 3/4\" (19.1mm)",  peso_unit: 2.235, unidad:"metro"},
  {codigo:"BCON-1",      descripcion:"Barra Construcción G60 1\" (25.4mm)",    peso_unit: 3.973, unidad:"metro"},
  {codigo:"BCON-1-3/8",  descripcion:"Barra Construcción G60 1 3/8\" (35.8mm)",peso_unit: 7.907, unidad:"metro"},
],

"BARRAS DE CONSTRUCCION ASTM A-706 GRADO 60 (SOLDABLES)": [
  {codigo:"BSOLD-5/8",   descripcion:"Barra Construcción Soldable G60 5/8\" (15.9mm)", peso_unit: 1.552, unidad:"metro"},
  {codigo:"BSOLD-3/4",   descripcion:"Barra Construcción Soldable G60 3/4\" (19.1mm)", peso_unit: 2.235, unidad:"metro"},
  {codigo:"BSOLD-1",     descripcion:"Barra Construcción Soldable G60 1\" (25.4mm)",   peso_unit: 3.973, unidad:"metro"},
],

"TUBERIA ACERO SIN COSTURA SCH 40/80/120/160 (ASTM A53-B / A106-B / API 5L-B)": [
  {codigo:"TUB-3/8-STD",  descripcion:"Tubería S/C NPS 3/8\" SCH40 STD (17.1mm DE, 2.31mm esp)",  peso_unit:  0.84, unidad:"metro"},
  {codigo:"TUB-3/8-XS",   descripcion:"Tubería S/C NPS 3/8\" SCH80 XS  (17.1mm DE, 3.20mm esp)",  peso_unit:  1.10, unidad:"metro"},
  {codigo:"TUB-1/2-STD",  descripcion:"Tubería S/C NPS 1/2\" SCH40 STD (21.3mm DE, 2.77mm esp)",  peso_unit:  1.27, unidad:"metro"},
  {codigo:"TUB-1/2-XS",   descripcion:"Tubería S/C NPS 1/2\" SCH80 XS  (21.3mm DE, 3.73mm esp)",  peso_unit:  1.62, unidad:"metro"},
  {codigo:"TUB-1/2-160",  descripcion:"Tubería S/C NPS 1/2\" SCH160    (21.3mm DE, 4.78mm esp)",  peso_unit:  1.95, unidad:"metro"},
  {codigo:"TUB-1/2-XXS",  descripcion:"Tubería S/C NPS 1/2\" XXS       (21.3mm DE, 7.47mm esp)",  peso_unit:  2.55, unidad:"metro"},
  {codigo:"TUB-3/4-STD",  descripcion:"Tubería S/C NPS 3/4\" SCH40 STD (26.7mm DE, 2.87mm esp)",  peso_unit:  1.69, unidad:"metro"},
  {codigo:"TUB-3/4-XS",   descripcion:"Tubería S/C NPS 3/4\" SCH80 XS  (26.7mm DE, 3.91mm esp)",  peso_unit:  2.20, unidad:"metro"},
  {codigo:"TUB-3/4-160",  descripcion:"Tubería S/C NPS 3/4\" SCH160    (26.7mm DE, 5.56mm esp)",  peso_unit:  2.90, unidad:"metro"},
  {codigo:"TUB-3/4-XXS",  descripcion:"Tubería S/C NPS 3/4\" XXS       (26.7mm DE, 7.82mm esp)",  peso_unit:  3.64, unidad:"metro"},
  {codigo:"TUB-1-STD",    descripcion:"Tubería S/C NPS 1\" SCH40 STD   (33.4mm DE, 3.38mm esp)",  peso_unit:  2.50, unidad:"metro"},
  {codigo:"TUB-1-XS",     descripcion:"Tubería S/C NPS 1\" SCH80 XS    (33.4mm DE, 4.55mm esp)",  peso_unit:  3.24, unidad:"metro"},
  {codigo:"TUB-1-160",    descripcion:"Tubería S/C NPS 1\" SCH160      (33.4mm DE, 6.35mm esp)",  peso_unit:  4.24, unidad:"metro"},
  {codigo:"TUB-1-XXS",    descripcion:"Tubería S/C NPS 1\" XXS         (33.4mm DE, 9.09mm esp)",  peso_unit:  5.45, unidad:"metro"},
  {codigo:"TUB-1.25-STD", descripcion:"Tubería S/C NPS 1 1/4\" SCH40 STD (42.2mm DE, 3.56mm esp)",peso_unit:  3.39, unidad:"metro"},
  {codigo:"TUB-1.25-XS",  descripcion:"Tubería S/C NPS 1 1/4\" SCH80 XS  (42.2mm DE, 4.85mm esp)",peso_unit:  4.47, unidad:"metro"},
  {codigo:"TUB-1.25-160", descripcion:"Tubería S/C NPS 1 1/4\" SCH160    (42.2mm DE, 6.35mm esp)",peso_unit:  5.61, unidad:"metro"},
  {codigo:"TUB-1.25-XXS", descripcion:"Tubería S/C NPS 1 1/4\" XXS       (42.2mm DE, 9.70mm esp)",peso_unit:  7.77, unidad:"metro"},
  {codigo:"TUB-1.5-STD",  descripcion:"Tubería S/C NPS 1 1/2\" SCH40 STD (48.3mm DE, 3.68mm esp)",peso_unit:  4.05, unidad:"metro"},
  {codigo:"TUB-1.5-XS",   descripcion:"Tubería S/C NPS 1 1/2\" SCH80 XS  (48.3mm DE, 5.08mm esp)",peso_unit:  5.41, unidad:"metro"},
  {codigo:"TUB-1.5-160",  descripcion:"Tubería S/C NPS 1 1/2\" SCH160    (48.3mm DE, 7.14mm esp)",peso_unit:  7.25, unidad:"metro"},
  {codigo:"TUB-1.5-XXS",  descripcion:"Tubería S/C NPS 1 1/2\" XXS       (48.3mm DE, 10.16mm esp)",peso_unit: 9.56, unidad:"metro"},
  {codigo:"TUB-2-STD",    descripcion:"Tubería S/C NPS 2\" SCH40 STD   (60.3mm DE, 3.91mm esp)",  peso_unit:  5.44, unidad:"metro"},
  {codigo:"TUB-2-XS",     descripcion:"Tubería S/C NPS 2\" SCH80 XS    (60.3mm DE, 5.54mm esp)",  peso_unit:  7.48, unidad:"metro"},
  {codigo:"TUB-2-160",    descripcion:"Tubería S/C NPS 2\" SCH160      (60.3mm DE, 8.74mm esp)",  peso_unit: 11.11, unidad:"metro"},
  {codigo:"TUB-2-XXS",    descripcion:"Tubería S/C NPS 2\" XXS         (60.3mm DE, 11.07mm esp)", peso_unit: 13.44, unidad:"metro"},
  {codigo:"TUB-2.5-STD",  descripcion:"Tubería S/C NPS 2 1/2\" SCH40 STD (73.0mm DE, 5.16mm esp)",peso_unit:  8.63, unidad:"metro"},
  {codigo:"TUB-2.5-XS",   descripcion:"Tubería S/C NPS 2 1/2\" SCH80 XS  (73.0mm DE, 7.01mm esp)",peso_unit: 11.41, unidad:"metro"},
  {codigo:"TUB-2.5-160",  descripcion:"Tubería S/C NPS 2 1/2\" SCH160    (73.0mm DE, 9.53mm esp)",peso_unit: 14.92, unidad:"metro"},
  {codigo:"TUB-2.5-XXS",  descripcion:"Tubería S/C NPS 2 1/2\" XXS       (73.0mm DE, 14.02mm esp)",peso_unit:20.29, unidad:"metro"},
  {codigo:"TUB-3-STD",    descripcion:"Tubería S/C NPS 3\" SCH40 STD   (88.9mm DE, 5.49mm esp)",  peso_unit: 11.29, unidad:"metro"},
  {codigo:"TUB-3-XS",     descripcion:"Tubería S/C NPS 3\" SCH80 XS    (88.9mm DE, 7.62mm esp)",  peso_unit: 15.27, unidad:"metro"},
  {codigo:"TUB-3-160",    descripcion:"Tubería S/C NPS 3\" SCH160      (88.9mm DE, 11.13mm esp)", peso_unit: 21.35, unidad:"metro"},
  {codigo:"TUB-3-XXS",    descripcion:"Tubería S/C NPS 3\" XXS         (88.9mm DE, 15.24mm esp)", peso_unit: 27.68, unidad:"metro"},
  {codigo:"TUB-3.5-STD",  descripcion:"Tubería S/C NPS 3 1/2\" SCH40 STD (101.6mm DE, 5.74mm esp)",peso_unit: 13.57, unidad:"metro"},
  {codigo:"TUB-3.5-XS",   descripcion:"Tubería S/C NPS 3 1/2\" SCH80 XS  (101.6mm DE, 8.08mm esp)",peso_unit: 18.63, unidad:"metro"},
  {codigo:"TUB-4-STD",    descripcion:"Tubería S/C NPS 4\" SCH40 STD   (114.3mm DE, 6.02mm esp)", peso_unit: 16.07, unidad:"metro"},
  {codigo:"TUB-4-XS",     descripcion:"Tubería S/C NPS 4\" SCH80 XS    (114.3mm DE, 8.56mm esp)", peso_unit: 22.32, unidad:"metro"},
  {codigo:"TUB-4-120",    descripcion:"Tubería S/C NPS 4\" SCH120      (114.3mm DE, 11.13mm esp)",peso_unit: 28.32, unidad:"metro"},
  {codigo:"TUB-4-160",    descripcion:"Tubería S/C NPS 4\" SCH160      (114.3mm DE, 13.49mm esp)",peso_unit: 33.54, unidad:"metro"},
  {codigo:"TUB-4-XXS",    descripcion:"Tubería S/C NPS 4\" XXS         (114.3mm DE, 17.12mm esp)",peso_unit: 41.03, unidad:"metro"},
  {codigo:"TUB-5-STD",    descripcion:"Tubería S/C NPS 5\" SCH40 STD   (141.3mm DE, 6.55mm esp)", peso_unit: 21.77, unidad:"metro"},
  {codigo:"TUB-5-XS",     descripcion:"Tubería S/C NPS 5\" SCH80 XS    (141.3mm DE, 9.53mm esp)", peso_unit: 30.94, unidad:"metro"},
  {codigo:"TUB-5-120",    descripcion:"Tubería S/C NPS 5\" SCH120      (141.3mm DE, 12.70mm esp)",peso_unit: 40.28, unidad:"metro"},
  {codigo:"TUB-5-160",    descripcion:"Tubería S/C NPS 5\" SCH160      (141.3mm DE, 15.88mm esp)",peso_unit: 49.11, unidad:"metro"},
  {codigo:"TUB-5-XXS",    descripcion:"Tubería S/C NPS 5\" XXS         (141.3mm DE, 19.05mm esp)",peso_unit: 57.43, unidad:"metro"},
  {codigo:"TUB-6-STD",    descripcion:"Tubería S/C NPS 6\" SCH40 STD   (168.3mm DE, 7.11mm esp)", peso_unit: 28.26, unidad:"metro"},
  {codigo:"TUB-6-XS",     descripcion:"Tubería S/C NPS 6\" SCH80 XS    (168.3mm DE, 10.97mm esp)",peso_unit: 42.56, unidad:"metro"},
  {codigo:"TUB-6-120",    descripcion:"Tubería S/C NPS 6\" SCH120      (168.3mm DE, 14.27mm esp)",peso_unit: 54.20, unidad:"metro"},
  {codigo:"TUB-6-160",    descripcion:"Tubería S/C NPS 6\" SCH160      (168.3mm DE, 18.26mm esp)",peso_unit: 67.56, unidad:"metro"},
  {codigo:"TUB-6-XXS",    descripcion:"Tubería S/C NPS 6\" XXS         (168.3mm DE, 21.95mm esp)",peso_unit: 79.22, unidad:"metro"},
  {codigo:"TUB-8-STD",    descripcion:"Tubería S/C NPS 8\" SCH40 STD   (219.1mm DE, 8.18mm esp)", peso_unit: 42.55, unidad:"metro"},
  {codigo:"TUB-8-XS",     descripcion:"Tubería S/C NPS 8\" SCH80 XS    (219.1mm DE, 12.70mm esp)",peso_unit: 64.64, unidad:"metro"},
  {codigo:"TUB-8-100",    descripcion:"Tubería S/C NPS 8\" SCH100      (219.1mm DE, 15.09mm esp)",peso_unit: 75.92, unidad:"metro"},
  {codigo:"TUB-8-120",    descripcion:"Tubería S/C NPS 8\" SCH120      (219.1mm DE, 18.26mm esp)",peso_unit: 90.44, unidad:"metro"},
  {codigo:"TUB-8-140",    descripcion:"Tubería S/C NPS 8\" SCH140      (219.1mm DE, 20.62mm esp)",peso_unit:100.92, unidad:"metro"},
  {codigo:"TUB-8-XXS",    descripcion:"Tubería S/C NPS 8\" XXS         (219.1mm DE, 22.23mm esp)",peso_unit:107.88, unidad:"metro"},
  {codigo:"TUB-8-160",    descripcion:"Tubería S/C NPS 8\" SCH160      (219.1mm DE, 23.01mm esp)",peso_unit:172.21, unidad:"metro"},
  {codigo:"TUB-10-STD",   descripcion:"Tubería S/C NPS 10\" SCH40 STD  (273.0mm DE, 9.27mm esp)", peso_unit: 60.29, unidad:"metro"},
  {codigo:"TUB-10-XS",    descripcion:"Tubería S/C NPS 10\" SCH80 XS   (273.0mm DE, 12.70mm esp)",peso_unit: 81.52, unidad:"metro"},
  {codigo:"TUB-10-100",   descripcion:"Tubería S/C NPS 10\" SCH100     (273.0mm DE, 15.09mm esp)",peso_unit: 95.97, unidad:"metro"},
  {codigo:"TUB-10-120",   descripcion:"Tubería S/C NPS 10\" SCH120     (273.0mm DE, 21.44mm esp)",peso_unit:133.00, unidad:"metro"},
  {codigo:"TUB-10-XXS",   descripcion:"Tubería S/C NPS 10\" XXS        (273.0mm DE, 25.40mm esp)",peso_unit:155.09, unidad:"metro"},
  {codigo:"TUB-10-160",   descripcion:"Tubería S/C NPS 10\" SCH160     (273.0mm DE, 28.58mm esp)",peso_unit:172.21, unidad:"metro"},
  {codigo:"TUB-12-STD",   descripcion:"Tubería S/C NPS 12\" SCH40 STD  (323.8mm DE, 9.53mm esp)", peso_unit: 73.78, unidad:"metro"},
  {codigo:"TUB-12-40",    descripcion:"Tubería S/C NPS 12\" SCH40      (323.8mm DE, 10.31mm esp)",peso_unit: 79.70, unidad:"metro"},
  {codigo:"TUB-12-XS",    descripcion:"Tubería S/C NPS 12\" SCH80 XS   (323.8mm DE, 12.70mm esp)",peso_unit: 97.43, unidad:"metro"},
  {codigo:"TUB-12-60",    descripcion:"Tubería S/C NPS 12\" SCH60      (323.8mm DE, 14.27mm esp)",peso_unit:108.92, unidad:"metro"},
  {codigo:"TUB-12-80",    descripcion:"Tubería S/C NPS 12\" SCH80      (323.8mm DE, 17.48mm esp)",peso_unit:132.04, unidad:"metro"},
  {codigo:"TUB-14-STD",   descripcion:"Tubería S/C NPS 14\" SCH30 STD  (355.6mm DE, 9.53mm esp)", peso_unit: 81.33, unidad:"metro"},
  {codigo:"TUB-14-40",    descripcion:"Tubería S/C NPS 14\" SCH40      (355.6mm DE, 11.13mm esp)",peso_unit: 94.55, unidad:"metro"},
  {codigo:"TUB-14-XS",    descripcion:"Tubería S/C NPS 14\" SCH80 XS   (355.6mm DE, 15.09mm esp)",peso_unit:126.71, unidad:"metro"},
  {codigo:"TUB-14-60",    descripcion:"Tubería S/C NPS 14\" SCH60      (355.6mm DE, 19.05mm esp)",peso_unit:158.10, unidad:"metro"},
],

"TUBOS ACERO SOLDADOS ERW (ARCO)": [
  {codigo:"ERW-16x9.5",   descripcion:"Tubo ERW Ø 406.4 mm (16\") × esp. 9.5 mm (3/8\")",  peso_unit:  92.98, unidad:"metro"},
  {codigo:"ERW-16x12.7",  descripcion:"Tubo ERW Ø 406.4 mm (16\") × esp. 12.7 mm (1/2\")", peso_unit: 123.30, unidad:"metro"},
  {codigo:"ERW-18x9.5",   descripcion:"Tubo ERW Ø 457.0 mm (18\") × esp. 9.5 mm (3/8\")",  peso_unit: 104.84, unidad:"metro"},
  {codigo:"ERW-18x12.7",  descripcion:"Tubo ERW Ø 457.0 mm (18\") × esp. 12.7 mm (1/2\")", peso_unit: 139.15, unidad:"metro"},
  {codigo:"ERW-20x9.5",   descripcion:"Tubo ERW Ø 508.0 mm (20\") × esp. 9.5 mm (3/8\")",  peso_unit: 116.78, unidad:"metro"},
  {codigo:"ERW-20x12.7",  descripcion:"Tubo ERW Ø 508.0 mm (20\") × esp. 12.7 mm (1/2\")", peso_unit: 155.12, unidad:"metro"},
  {codigo:"ERW-22x9.5",   descripcion:"Tubo ERW Ø 559.0 mm (22\") × esp. 9.5 mm (3/8\")",  peso_unit: 128.73, unidad:"metro"},
  {codigo:"ERW-22x12.7",  descripcion:"Tubo ERW Ø 559.0 mm (22\") × esp. 12.7 mm (1/2\")", peso_unit: 171.09, unidad:"metro"},
  {codigo:"ERW-24x9.5",   descripcion:"Tubo ERW Ø 610.0 mm (24\") × esp. 9.5 mm (3/8\")",  peso_unit: 140.68, unidad:"metro"},
  {codigo:"ERW-24x12.7",  descripcion:"Tubo ERW Ø 610.0 mm (24\") × esp. 12.7 mm (1/2\")", peso_unit: 187.06, unidad:"metro"},
  {codigo:"ERW-30x9.5",   descripcion:"Tubo ERW Ø 762.0 mm (30\") × esp. 9.5 mm (3/8\")",  peso_unit: 176.29, unidad:"metro"},
  {codigo:"ERW-30x12.7",  descripcion:"Tubo ERW Ø 762.0 mm (30\") × esp. 12.7 mm (1/2\")", peso_unit: 234.67, unidad:"metro"},
],

"TUBOS REDONDOS LAC (SAE 1010 / ISO 65)": [
  {codigo:"TUB-LAC-1/4x1.80", descripcion:"Tubo LAC DN 1/4\" Ø13.7 mm × esp. 1.80 mm", peso_unit: 0.515, unidad:"metro"},
  {codigo:"TUB-LAC-1/4x2.00", descripcion:"Tubo LAC DN 1/4\" Ø13.7 mm × esp. 2.00 mm", peso_unit: 0.570, unidad:"metro"},
  {codigo:"TUB-LAC-3/8x1.80", descripcion:"Tubo LAC DN 3/8\" Ø17.2 mm × esp. 1.80 mm", peso_unit: 0.670, unidad:"metro"},
  {codigo:"TUB-LAC-3/8x2.00", descripcion:"Tubo LAC DN 3/8\" Ø17.2 mm × esp. 2.00 mm", peso_unit: 0.742, unidad:"metro"},
  {codigo:"TUB-LAC-1/2x1.80", descripcion:"Tubo LAC DN 1/2\" Ø21.3 mm × esp. 1.80 mm", peso_unit: 0.910, unidad:"metro"},
  {codigo:"TUB-LAC-1/2x2.00", descripcion:"Tubo LAC DN 1/2\" Ø21.3 mm × esp. 2.00 mm", peso_unit: 0.947, unidad:"metro"},
  {codigo:"TUB-LAC-1/2x2.30", descripcion:"Tubo LAC DN 1/2\" Ø21.3 mm × esp. 2.30 mm", peso_unit: 1.100, unidad:"metro"},
  {codigo:"TUB-LAC-1/2x2.35", descripcion:"Tubo LAC DN 1/2\" Ø21.3 mm × esp. 2.35 mm", peso_unit: 1.100, unidad:"metro"},
  {codigo:"TUB-LAC-1/2x2.50", descripcion:"Tubo LAC DN 1/2\" Ø21.3 mm × esp. 2.50 mm", peso_unit: 1.280, unidad:"metro"},
  {codigo:"TUB-LAC-1/2x2.65", descripcion:"Tubo LAC DN 1/2\" Ø21.3 mm × esp. 2.65 mm", peso_unit: 1.356, unidad:"metro"},
  {codigo:"TUB-LAC-3/4x1.80", descripcion:"Tubo LAC DN 3/4\" Ø26.9 mm × esp. 1.80 mm", peso_unit: 1.150, unidad:"metro"},
  {codigo:"TUB-LAC-3/4x2.00", descripcion:"Tubo LAC DN 3/4\" Ø26.9 mm × esp. 2.00 mm", peso_unit: 1.240, unidad:"metro"},
  {codigo:"TUB-LAC-3/4x2.30", descripcion:"Tubo LAC DN 3/4\" Ø26.9 mm × esp. 2.30 mm", peso_unit: 1.410, unidad:"metro"},
  {codigo:"TUB-LAC-3/4x2.35", descripcion:"Tubo LAC DN 3/4\" Ø26.9 mm × esp. 2.35 mm", peso_unit: 1.426, unidad:"metro"},
  {codigo:"TUB-LAC-3/4x2.50", descripcion:"Tubo LAC DN 3/4\" Ø26.9 mm × esp. 2.50 mm", peso_unit: 1.580, unidad:"metro"},
  {codigo:"TUB-LAC-3/4x2.65", descripcion:"Tubo LAC DN 3/4\" Ø26.9 mm × esp. 2.65 mm", peso_unit: 1.560, unidad:"metro"},
  {codigo:"TUB-LAC-1x2.00",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 2.00 mm",   peso_unit: 1.620, unidad:"metro"},
  {codigo:"TUB-LAC-1x2.30",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 2.30 mm",   peso_unit: 1.800, unidad:"metro"},
  {codigo:"TUB-LAC-1x2.50",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 2.50 mm",   peso_unit: 1.877, unidad:"metro"},
  {codigo:"TUB-LAC-1x2.65",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 2.65 mm",   peso_unit: 1.980, unidad:"metro"},
  {codigo:"TUB-LAC-1x2.90",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 2.90 mm",   peso_unit: 2.200, unidad:"metro"},
  {codigo:"TUB-LAC-1x3.00",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 3.00 mm",   peso_unit: 2.280, unidad:"metro"},
  {codigo:"TUB-LAC-1x3.25",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 3.25 mm",   peso_unit: 2.410, unidad:"metro"},
  {codigo:"TUB-LAC-1x3.30",   descripcion:"Tubo LAC DN 1\" Ø33.7 mm × esp. 3.30 mm",   peso_unit: 2.500, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x2.00", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 2.00 mm", peso_unit: 2.040, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x2.30", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 2.30 mm", peso_unit: 2.340, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x2.50", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 2.50 mm", peso_unit: 2.340, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x2.65", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 2.65 mm", peso_unit: 2.540, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x2.90", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 2.90 mm", peso_unit: 2.820, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x3.00", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 3.00 mm", peso_unit: 2.840, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x3.25", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 3.25 mm", peso_unit: 3.100, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/4x3.30", descripcion:"Tubo LAC DN 1 1/4\" Ø42.4 mm × esp. 3.30 mm", peso_unit: 3.140, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x2.00", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 2.00 mm", peso_unit: 2.350, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x2.30", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 2.30 mm", peso_unit: 2.630, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x2.50", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 2.50 mm", peso_unit: 2.840, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x2.90", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 2.90 mm", peso_unit: 3.230, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x3.00", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 3.00 mm", peso_unit: 3.370, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x3.25", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 3.25 mm", peso_unit: 3.560, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x3.30", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 3.30 mm", peso_unit: 3.800, unidad:"metro"},
  {codigo:"TUB-LAC-1-1/2x4.00", descripcion:"Tubo LAC DN 1 1/2\" Ø48.3 mm × esp. 4.00 mm", peso_unit: 4.558, unidad:"metro"},
  {codigo:"TUB-LAC-2x2.00",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 2.00 mm",   peso_unit: 2.920, unidad:"metro"},
  {codigo:"TUB-LAC-2x2.30",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 2.30 mm",   peso_unit: 3.300, unidad:"metro"},
  {codigo:"TUB-LAC-2x2.50",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 2.50 mm",   peso_unit: 3.630, unidad:"metro"},
  {codigo:"TUB-LAC-2x2.90",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 2.90 mm",   peso_unit: 4.080, unidad:"metro"},
  {codigo:"TUB-LAC-2x3.00",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 3.00 mm",   peso_unit: 4.230, unidad:"metro"},
  {codigo:"TUB-LAC-2x3.25",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 3.25 mm",   peso_unit: 4.490, unidad:"metro"},
  {codigo:"TUB-LAC-2x3.30",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 3.30 mm",   peso_unit: 4.750, unidad:"metro"},
  {codigo:"TUB-LAC-2x3.60",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 3.60 mm",   peso_unit: 5.030, unidad:"metro"},
  {codigo:"TUB-LAC-2x4.00",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 4.00 mm",   peso_unit: 5.630, unidad:"metro"},
  {codigo:"TUB-LAC-2x4.50",   descripcion:"Tubo LAC DN 2\" Ø60.3 mm × esp. 4.50 mm",   peso_unit: 6.281, unidad:"metro"},
  {codigo:"TUB-LAC-2-1/2x2.50", descripcion:"Tubo LAC DN 2 1/2\" Ø73.0 mm × esp. 2.50 mm", peso_unit: 4.590, unidad:"metro"},
  {codigo:"TUB-LAC-2-1/2x3.00", descripcion:"Tubo LAC DN 2 1/2\" Ø73.0 mm × esp. 3.00 mm", peso_unit: 5.490, unidad:"metro"},
  {codigo:"TUB-LAC-2-1/2x3.30", descripcion:"Tubo LAC DN 2 1/2\" Ø73.0 mm × esp. 3.30 mm", peso_unit: 6.020, unidad:"metro"},
  {codigo:"TUB-LAC-2-1/2x4.00", descripcion:"Tubo LAC DN 2 1/2\" Ø73.0 mm × esp. 4.00 mm", peso_unit: 7.260, unidad:"metro"},
  {codigo:"TUB-LAC-2-1/2x4.50", descripcion:"Tubo LAC DN 2 1/2\" Ø73.0 mm × esp. 4.50 mm", peso_unit: 8.114, unidad:"metro"},
  {codigo:"TUB-LAC-3x2.50",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 2.50 mm",   peso_unit:  5.380, unidad:"metro"},
  {codigo:"TUB-LAC-3x3.00",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 3.00 mm",   peso_unit:  6.440, unidad:"metro"},
  {codigo:"TUB-LAC-3x3.20",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 3.20 mm",   peso_unit:  6.720, unidad:"metro"},
  {codigo:"TUB-LAC-3x3.30",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 3.30 mm",   peso_unit:  7.060, unidad:"metro"},
  {codigo:"TUB-LAC-3x3.60",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 3.60 mm",   peso_unit:  7.550, unidad:"metro"},
  {codigo:"TUB-LAC-3x4.00",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 4.00 mm",   peso_unit:  8.360, unidad:"metro"},
  {codigo:"TUB-LAC-3x4.50",   descripcion:"Tubo LAC DN 3\" Ø88.9 mm × esp. 4.50 mm",   peso_unit:  9.535, unidad:"metro"},
  {codigo:"TUB-LAC-4x3.00",   descripcion:"Tubo LAC DN 4\" Ø114.3 mm × esp. 3.00 mm",  peso_unit:  8.320, unidad:"metro"},
  {codigo:"TUB-LAC-4x3.30",   descripcion:"Tubo LAC DN 4\" Ø114.3 mm × esp. 3.30 mm",  peso_unit:  9.130, unidad:"metro"},
  {codigo:"TUB-LAC-4x3.60",   descripcion:"Tubo LAC DN 4\" Ø114.3 mm × esp. 3.60 mm",  peso_unit:  9.750, unidad:"metro"},
  {codigo:"TUB-LAC-4x4.00",   descripcion:"Tubo LAC DN 4\" Ø114.3 mm × esp. 4.00 mm",  peso_unit: 10.800, unidad:"metro"},
  {codigo:"TUB-LAC-4x4.50",   descripcion:"Tubo LAC DN 4\" Ø114.3 mm × esp. 4.50 mm",  peso_unit: 12.200, unidad:"metro"},
],

"TUBOS CUADRADOS LAC (COPANT 518 - TIPO I)": [
  {codigo:"TCUAD-1x2.0",      descripcion:"Tubo Cuadrado LAC 1\" (25.40mm) e=2.0mm",    peso_unit:  9.212, unidad:"mm"},
  {codigo:"TCUAD-1x2.3",      descripcion:"Tubo Cuadrado LAC 1\" (25.40mm) e=2.3mm",    peso_unit: 10.539, unidad:"mm"},
  {codigo:"TCUAD-1x2.5",      descripcion:"Tubo Cuadrado LAC 1\" (25.40mm) e=2.5mm",    peso_unit: 11.215, unidad:"mm"},
  {codigo:"TCUAD-1x3.0",      descripcion:"Tubo Cuadrado LAC 1\" (25.40mm) e=3.0mm",    peso_unit: 13.506, unidad:"mm"},
  {codigo:"TCUAD-1.25x2.0",   descripcion:"Tubo Cuadrado LAC 1 1/4\" (31.75mm) e=2.0mm",peso_unit: 12.047, unidad:"mm"},
  {codigo:"TCUAD-1.25x2.3",   descripcion:"Tubo Cuadrado LAC 1 1/4\" (31.75mm) e=2.3mm",peso_unit: 13.742, unidad:"mm"},
  {codigo:"TCUAD-1.25x2.5",   descripcion:"Tubo Cuadrado LAC 1 1/4\" (31.75mm) e=2.5mm",peso_unit: 14.974, unidad:"mm"},
  {codigo:"TCUAD-1.25x3.0",   descripcion:"Tubo Cuadrado LAC 1 1/4\" (31.75mm) e=3.0mm",peso_unit: 17.898, unidad:"mm"},
  {codigo:"TCUAD-1.5x2.0",    descripcion:"Tubo Cuadrado LAC 1 1/2\" (38.10mm) e=2.0mm",peso_unit: 14.188, unidad:"mm"},
  {codigo:"TCUAD-1.5x2.3",    descripcion:"Tubo Cuadrado LAC 1 1/2\" (38.10mm) e=2.3mm",peso_unit: 16.358, unidad:"mm"},
  {codigo:"TCUAD-1.5x2.5",    descripcion:"Tubo Cuadrado LAC 1 1/2\" (38.10mm) e=2.5mm",peso_unit: 17.675, unidad:"mm"},
  {codigo:"TCUAD-1.5x3.0",    descripcion:"Tubo Cuadrado LAC 1 1/2\" (38.10mm) e=3.0mm",peso_unit: 21.175, unidad:"mm"},
  {codigo:"TCUAD-1.5x4.0",    descripcion:"Tubo Cuadrado LAC 1 1/2\" (38.10mm) e=4.0mm",peso_unit: 27.951, unidad:"mm"},
  {codigo:"TCUAD-2x2.0",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=2.0mm",    peso_unit: 18.651, unidad:"mm"},
  {codigo:"TCUAD-2x2.3",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=2.3mm",    peso_unit: 21.559, unidad:"mm"},
  {codigo:"TCUAD-2x2.5",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=2.5mm",    peso_unit: 22.960, unidad:"mm"},
  {codigo:"TCUAD-2x3.0",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=3.0mm",    peso_unit: 27.411, unidad:"mm"},
  {codigo:"TCUAD-2x3.3",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=3.3mm",    peso_unit: 30.027, unidad:"mm"},
  {codigo:"TCUAD-2x4.0",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=4.0mm",    peso_unit: 36.009, unidad:"mm"},
  {codigo:"TCUAD-2x4.5",      descripcion:"Tubo Cuadrado LAC 2\" (50.80mm) e=4.5mm",    peso_unit: 40.192, unidad:"mm"},
  {codigo:"TCUAD-3x2.0",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=2.0mm",    peso_unit: 28.188, unidad:"mm"},
  {codigo:"TCUAD-3x2.3",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=2.3mm",    peso_unit: 32.499, unidad:"mm"},
  {codigo:"TCUAD-3x2.5",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=2.5mm",    peso_unit: 35.207, unidad:"mm"},
  {codigo:"TCUAD-3x3.0",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=3.0mm",    peso_unit: 42.195, unidad:"mm"},
  {codigo:"TCUAD-3x3.3",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=3.3mm",    peso_unit: 46.163, unidad:"mm"},
  {codigo:"TCUAD-3x4.0",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=4.0mm",    peso_unit: 55.672, unidad:"mm"},
  {codigo:"TCUAD-3x4.5",      descripcion:"Tubo Cuadrado LAC 3\" (76.20mm) e=4.5mm",    peso_unit: 62.313, unidad:"mm"},
  {codigo:"TCUAD-4x2.5",      descripcion:"Tubo Cuadrado LAC 4\" (101.60mm) e=2.5mm",   peso_unit: 47.282, unidad:"mm"},
  {codigo:"TCUAD-4x3.0",      descripcion:"Tubo Cuadrado LAC 4\" (101.60mm) e=3.0mm",   peso_unit: 56.527, unidad:"mm"},
  {codigo:"TCUAD-4x3.3",      descripcion:"Tubo Cuadrado LAC 4\" (101.60mm) e=3.3mm",   peso_unit: 61.993, unidad:"mm"},
  {codigo:"TCUAD-4x4.0",      descripcion:"Tubo Cuadrado LAC 4\" (101.60mm) e=4.0mm",   peso_unit: 74.766, unidad:"mm"},
  {codigo:"TCUAD-4x4.5",      descripcion:"Tubo Cuadrado LAC 4\" (101.60mm) e=4.5mm",   peso_unit: 83.794, unidad:"mm"},
  {codigo:"TCUAD-40mmx2.0",   descripcion:"Tubo Cuadrado LAC 40mm e=2.0mm",              peso_unit: 14.930, unidad:"mm"},
  {codigo:"TCUAD-40mmx2.5",   descripcion:"Tubo Cuadrado LAC 40mm e=2.5mm",              peso_unit: 18.262, unidad:"mm"},
  {codigo:"TCUAD-40mmx3.0",   descripcion:"Tubo Cuadrado LAC 40mm e=3.0mm",              peso_unit: 21.773, unidad:"mm"},
  {codigo:"TCUAD-40mmx4.0",   descripcion:"Tubo Cuadrado LAC 40mm e=4.0mm",              peso_unit: 28.302, unidad:"mm"},
  {codigo:"TCUAD-40mmx4.5",   descripcion:"Tubo Cuadrado LAC 40mm e=4.5mm",              peso_unit: 31.628, unidad:"mm"},
  {codigo:"TCUAD-50mmx2.0",   descripcion:"Tubo Cuadrado LAC 50mm e=2.0mm",              peso_unit: 18.591, unidad:"mm"},
  {codigo:"TCUAD-50mmx2.5",   descripcion:"Tubo Cuadrado LAC 50mm e=2.5mm",              peso_unit: 22.960, unidad:"mm"},
  {codigo:"TCUAD-50mmx3.0",   descripcion:"Tubo Cuadrado LAC 50mm e=3.0mm",              peso_unit: 27.411, unidad:"mm"},
  {codigo:"TCUAD-50mmx4.0",   descripcion:"Tubo Cuadrado LAC 50mm e=4.0mm",              peso_unit: 36.009, unidad:"mm"},
  {codigo:"TCUAD-50mmx4.5",   descripcion:"Tubo Cuadrado LAC 50mm e=4.5mm",              peso_unit: 40.086, unidad:"mm"},
  {codigo:"TCUAD-60mmx2.0",   descripcion:"Tubo Cuadrado LAC 60mm e=2.0mm",              peso_unit: 22.409, unidad:"mm"},
  {codigo:"TCUAD-60mmx2.5",   descripcion:"Tubo Cuadrado LAC 60mm e=2.5mm",              peso_unit: 27.658, unidad:"mm"},
  {codigo:"TCUAD-60mmx3.0",   descripcion:"Tubo Cuadrado LAC 60mm e=3.0mm",              peso_unit: 33.049, unidad:"mm"},
  {codigo:"TCUAD-60mmx4.0",   descripcion:"Tubo Cuadrado LAC 60mm e=4.0mm",              peso_unit: 43.050, unidad:"mm"},
  {codigo:"TCUAD-60mmx4.5",   descripcion:"Tubo Cuadrado LAC 60mm e=4.5mm",              peso_unit: 47.766, unidad:"mm"},
  {codigo:"TCUAD-100mmx2.5",  descripcion:"Tubo Cuadrado LAC 100mm e=2.5mm",             peso_unit: 47.160, unidad:"mm"},
  {codigo:"TCUAD-100mmx3.0",  descripcion:"Tubo Cuadrado LAC 100mm e=3.0mm",             peso_unit: 56.400, unidad:"mm"},
  {codigo:"TCUAD-100mmx4.0",  descripcion:"Tubo Cuadrado LAC 100mm e=4.0mm",             peso_unit: 74.400, unidad:"mm"},
  {codigo:"TCUAD-100mmx4.5",  descripcion:"Tubo Cuadrado LAC 100mm e=4.5mm",             peso_unit: 82.995, unidad:"mm"},
],

"TUBOS RECTANGULARES LAC (COPANT 518 - TIPO I)": [
  {codigo:"TRECT-1x2x2.0",    descripcion:"Tubo Rectangular LAC 1\"×2\" (25.4×50.8mm) e=2.0mm",  peso_unit: 14.188, unidad:"mm"},
  {codigo:"TRECT-2x3x2.5",    descripcion:"Tubo Rectangular LAC 2\"×3\" (50.8×76.2mm) e=2.5mm",  peso_unit: 29.280, unidad:"mm"},
  {codigo:"TRECT-2x3x3.0",    descripcion:"Tubo Rectangular LAC 2\"×3\" (50.8×76.2mm) e=3.0mm",  peso_unit: 34.980, unidad:"mm"},
  {codigo:"TRECT-2x3x4.0",    descripcion:"Tubo Rectangular LAC 2\"×3\" (50.8×76.2mm) e=4.0mm",  peso_unit: 45.724, unidad:"mm"},
  {codigo:"TRECT-2x3x4.5",    descripcion:"Tubo Rectangular LAC 2\"×3\" (50.8×76.2mm) e=4.5mm",  peso_unit: 47.901, unidad:"mm"},
  {codigo:"TRECT-2x4x2.5",    descripcion:"Tubo Rectangular LAC 2\"×4\" (50.8×101.6mm) e=2.5mm", peso_unit: 35.160, unidad:"mm"},
  {codigo:"TRECT-2x4x3.0",    descripcion:"Tubo Rectangular LAC 2\"×4\" (50.8×101.6mm) e=3.0mm", peso_unit: 42.100, unidad:"mm"},
  {codigo:"TRECT-2x4x4.0",    descripcion:"Tubo Rectangular LAC 2\"×4\" (50.8×101.6mm) e=4.0mm", peso_unit: 55.568, unidad:"mm"},
  {codigo:"TRECT-2x4x4.5",    descripcion:"Tubo Rectangular LAC 2\"×4\" (50.8×101.6mm) e=4.5mm", peso_unit: 62.196, unidad:"mm"},
  {codigo:"TRECT-40x50x2.0",  descripcion:"Tubo Rectangular LAC 40×50mm e=2.0mm",  peso_unit: 16.676, unidad:"mm"},
  {codigo:"TRECT-40x50x2.5",  descripcion:"Tubo Rectangular LAC 40×50mm e=2.5mm",  peso_unit: 20.712, unidad:"mm"},
  {codigo:"TRECT-40x50x3.0",  descripcion:"Tubo Rectangular LAC 40×50mm e=3.0mm",  peso_unit: 24.711, unidad:"mm"},
  {codigo:"TRECT-40x50x4.0",  descripcion:"Tubo Rectangular LAC 40×50mm e=4.0mm",  peso_unit: 31.966, unidad:"mm"},
  {codigo:"TRECT-40x50x4.5",  descripcion:"Tubo Rectangular LAC 40×50mm e=4.5mm",  peso_unit: 35.410, unidad:"mm"},
  {codigo:"TRECT-40x60x2.0",  descripcion:"Tubo Rectangular LAC 40×60mm e=2.0mm",  peso_unit: 18.591, unidad:"mm"},
  {codigo:"TRECT-40x60x2.5",  descripcion:"Tubo Rectangular LAC 40×60mm e=2.5mm",  peso_unit: 23.239, unidad:"mm"},
  {codigo:"TRECT-40x60x3.0",  descripcion:"Tubo Rectangular LAC 40×60mm e=3.0mm",  peso_unit: 27.743, unidad:"mm"},
  {codigo:"TRECT-40x60x4.0",  descripcion:"Tubo Rectangular LAC 40×60mm e=4.0mm",  peso_unit: 36.009, unidad:"mm"},
  {codigo:"TRECT-40x60x4.5",  descripcion:"Tubo Rectangular LAC 40×60mm e=4.5mm",  peso_unit: 39.959, unidad:"mm"},
  {codigo:"TRECT-40x80x2.0",  descripcion:"Tubo Rectangular LAC 40×80mm e=2.0mm",  peso_unit: 22.407, unidad:"mm"},
  {codigo:"TRECT-40x80x2.5",  descripcion:"Tubo Rectangular LAC 40×80mm e=2.5mm",  peso_unit: 27.640, unidad:"mm"},
  {codigo:"TRECT-40x80x3.0",  descripcion:"Tubo Rectangular LAC 40×80mm e=3.0mm",  peso_unit: 33.024, unidad:"mm"},
  {codigo:"TRECT-40x80x4.0",  descripcion:"Tubo Rectangular LAC 40×80mm e=4.0mm",  peso_unit: 43.050, unidad:"mm"},
  {codigo:"TRECT-40x80x4.5",  descripcion:"Tubo Rectangular LAC 40×80mm e=4.5mm",  peso_unit: 47.880, unidad:"mm"},
  {codigo:"TRECT-50x70x2.0",  descripcion:"Tubo Rectangular LAC 50×70mm e=2.0mm",  peso_unit: 22.407, unidad:"mm"},
  {codigo:"TRECT-50x70x2.5",  descripcion:"Tubo Rectangular LAC 50×70mm e=2.5mm",  peso_unit: 27.640, unidad:"mm"},
  {codigo:"TRECT-50x70x3.0",  descripcion:"Tubo Rectangular LAC 50×70mm e=3.0mm",  peso_unit: 33.024, unidad:"mm"},
  {codigo:"TRECT-50x70x4.0",  descripcion:"Tubo Rectangular LAC 50×70mm e=4.0mm",  peso_unit: 43.050, unidad:"mm"},
  {codigo:"TRECT-50x70x4.5",  descripcion:"Tubo Rectangular LAC 50×70mm e=4.5mm",  peso_unit: 47.880, unidad:"mm"},
  {codigo:"TRECT-50x100x2.5", descripcion:"Tubo Rectangular LAC 50×100mm e=2.5mm", peso_unit: 34.784, unidad:"mm"},
  {codigo:"TRECT-50x100x3.0", descripcion:"Tubo Rectangular LAC 50×100mm e=3.0mm", peso_unit: 41.597, unidad:"mm"},
  {codigo:"TRECT-50x100x4.0", descripcion:"Tubo Rectangular LAC 50×100mm e=4.0mm", peso_unit: 54.685, unidad:"mm"},
  {codigo:"TRECT-50x100x4.5", descripcion:"Tubo Rectangular LAC 50×100mm e=4.5mm", peso_unit: 60.960, unidad:"mm"},
],

"TUBOS DE ACERO ISO 65 - LIVIANOS Serie II": [
  {codigo:"TISO-LIV-1/4",   descripcion:"Tubo ISO 65 Liviano 1/4\" DN (ext.13.6mm) e=1.80mm",  peso_unit: 0.517, unidad:"metro"},
  {codigo:"TISO-LIV-3/8",   descripcion:"Tubo ISO 65 Liviano 3/8\" DN (ext.17.1mm) e=1.80mm",  peso_unit: 0.674, unidad:"metro"},
  {codigo:"TISO-LIV-1/2",   descripcion:"Tubo ISO 65 Liviano 1/2\" DN (ext.21.4mm) e=2.00mm",  peso_unit: 0.952, unidad:"metro"},
  {codigo:"TISO-LIV-3/4",   descripcion:"Tubo ISO 65 Liviano 3/4\" DN (ext.26.9mm) e=2.00mm",  peso_unit: 1.200, unidad:"metro"},
  {codigo:"TISO-LIV-1",     descripcion:"Tubo ISO 65 Liviano 1\" DN (ext.33.8mm) e=2.65mm",    peso_unit: 2.010, unidad:"metro"},
  {codigo:"TISO-LIV-1-1/4", descripcion:"Tubo ISO 65 Liviano 1 1/4\" DN (ext.42.5mm) e=2.65mm",peso_unit: 2.580, unidad:"metro"},
  {codigo:"TISO-LIV-1-1/2", descripcion:"Tubo ISO 65 Liviano 1 1/2\" DN (ext.48.4mm) e=2.65mm",peso_unit: 2.970, unidad:"metro"},
  {codigo:"TISO-LIV-2",     descripcion:"Tubo ISO 65 Liviano 2\" DN (ext.60.2mm) e=2.90mm",    peso_unit: 4.110, unidad:"metro"},
  {codigo:"TISO-LIV-2-1/2", descripcion:"Tubo ISO 65 Liviano 2 1/2\" DN (ext.76.0mm) e=3.00mm",peso_unit: 5.371, unidad:"metro"},
  {codigo:"TISO-LIV-3",     descripcion:"Tubo ISO 65 Liviano 3\" DN (ext.88.7mm) e=3.25mm",    peso_unit: 6.810, unidad:"metro"},
  {codigo:"TISO-LIV-4",     descripcion:"Tubo ISO 65 Liviano 4\" DN (ext.113.9mm) e=3.65mm",   peso_unit: 9.890, unidad:"metro"},
],

"TUBOS DE ACERO ISO 65 - STANDARD Serie I": [
  {codigo:"TISO-STD-1/4",   descripcion:"Tubo ISO 65 Standard 1/4\" DN (ext.13.9mm) e=2.00mm",  peso_unit:  0.573, unidad:"metro"},
  {codigo:"TISO-STD-3/8",   descripcion:"Tubo ISO 65 Standard 3/8\" DN (ext.17.4mm) e=2.00mm",  peso_unit:  0.747, unidad:"metro"},
  {codigo:"TISO-STD-1/2",   descripcion:"Tubo ISO 65 Standard 1/2\" DN (ext.21.7mm) e=2.35mm",  peso_unit:  1.100, unidad:"metro"},
  {codigo:"TISO-STD-3/4",   descripcion:"Tubo ISO 65 Standard 3/4\" DN (ext.27.1mm) e=2.35mm",  peso_unit:  1.410, unidad:"metro"},
  {codigo:"TISO-STD-1",     descripcion:"Tubo ISO 65 Standard 1\" DN (ext.34.0mm) e=2.90mm",    peso_unit:  2.210, unidad:"metro"},
  {codigo:"TISO-STD-1-1/4", descripcion:"Tubo ISO 65 Standard 1 1/4\" DN (ext.42.7mm) e=2.90mm",peso_unit:  2.840, unidad:"metro"},
  {codigo:"TISO-STD-1-1/2", descripcion:"Tubo ISO 65 Standard 1 1/2\" DN (ext.48.6mm) e=2.90mm",peso_unit:  3.260, unidad:"metro"},
  {codigo:"TISO-STD-2",     descripcion:"Tubo ISO 65 Standard 2\" DN (ext.60.7mm) e=3.25mm",    peso_unit:  4.560, unidad:"metro"},
  {codigo:"TISO-STD-2-1/2", descripcion:"Tubo ISO 65 Standard 2 1/2\" DN (ext.76.3mm) e=3.25mm",peso_unit:  5.810, unidad:"metro"},
  {codigo:"TISO-STD-3",     descripcion:"Tubo ISO 65 Standard 3\" DN (ext.89.4mm) e=3.65mm",    peso_unit:  7.650, unidad:"metro"},
  {codigo:"TISO-STD-4",     descripcion:"Tubo ISO 65 Standard 4\" DN (ext.114.9mm) e=4.05mm",   peso_unit: 11.000, unidad:"metro"},
],

"TUBOS DE ACERO ISO 65 - SEMI PESADOS Serie Medium": [
  {codigo:"TISO-SMP-3/8",   descripcion:"Tubo ISO 65 Semi Pesado 3/8\" DN (ext.17.5mm) e=2.35mm",  peso_unit:  0.852, unidad:"metro"},
  {codigo:"TISO-SMP-1/2",   descripcion:"Tubo ISO 65 Semi Pesado 1/2\" DN (ext.21.8mm) e=2.65mm",  peso_unit:  1.220, unidad:"metro"},
  {codigo:"TISO-SMP-3/4",   descripcion:"Tubo ISO 65 Semi Pesado 3/4\" DN (ext.27.3mm) e=2.65mm",  peso_unit:  1.580, unidad:"metro"},
  {codigo:"TISO-SMP-1",     descripcion:"Tubo ISO 65 Semi Pesado 1\" DN (ext.34.2mm) e=3.25mm",    peso_unit:  2.440, unidad:"metro"},
  {codigo:"TISO-SMP-1-1/4", descripcion:"Tubo ISO 65 Semi Pesado 1 1/4\" DN (ext.42.9mm) e=3.25mm",peso_unit:  3.140, unidad:"metro"},
  {codigo:"TISO-SMP-1-1/2", descripcion:"Tubo ISO 65 Semi Pesado 1 1/2\" DN (ext.48.8mm) e=3.25mm",peso_unit:  3.610, unidad:"metro"},
  {codigo:"TISO-SMP-2",     descripcion:"Tubo ISO 65 Semi Pesado 2\" DN (ext.60.8mm) e=3.65mm",    peso_unit:  5.100, unidad:"metro"},
  {codigo:"TISO-SMP-2-1/2", descripcion:"Tubo ISO 65 Semi Pesado 2 1/2\" DN (ext.76.6mm) e=3.65mm",peso_unit:  6.510, unidad:"metro"},
  {codigo:"TISO-SMP-3",     descripcion:"Tubo ISO 65 Semi Pesado 3\" DN (ext.89.5mm) e=4.05mm",    peso_unit:  8.470, unidad:"metro"},
  {codigo:"TISO-SMP-4",     descripcion:"Tubo ISO 65 Semi Pesado 4\" DN (ext.115.0mm) e=4.50mm",   peso_unit: 12.100, unidad:"metro"},
  {codigo:"TISO-SMP-5",     descripcion:"Tubo ISO 65 Semi Pesado 5\" DN (ext.140.8mm) e=5.00mm",   peso_unit: 16.609, unidad:"metro"},
  {codigo:"TISO-SMP-6",     descripcion:"Tubo ISO 65 Semi Pesado 6\" DN (ext.166.5mm) e=5.00mm",   peso_unit: 19.200, unidad:"metro"},
],

"TUBOS ELECTRO SOLDADOS REDONDOS (SAE 1010 LAF - COPANT 518 TIPO II)": [
  {codigo:"TESR-3/8x0.9",    descripcion:"Tubo ES Redondo 3/8\" (9.53mm) e=0.9mm",   peso_unit:  0.835, unidad:"mm"},
  {codigo:"TESR-3/8x1.0",    descripcion:"Tubo ES Redondo 3/8\" (9.53mm) e=1.0mm",   peso_unit:  1.165, unidad:"mm"},
  {codigo:"TESR-3/8x1.2",    descripcion:"Tubo ES Redondo 3/8\" (9.53mm) e=1.2mm",   peso_unit:  1.393, unidad:"mm"},
  {codigo:"TESR-3/8x1.5",    descripcion:"Tubo ES Redondo 3/8\" (9.53mm) e=1.5mm",   peso_unit:  1.716, unidad:"mm"},
  {codigo:"TESR-3/8x2.0",    descripcion:"Tubo ES Redondo 3/8\" (9.53mm) e=2.0mm",   peso_unit:  2.229, unidad:"mm"},
  {codigo:"TESR-1/2x0.8",    descripcion:"Tubo ES Redondo 1/2\" (12.70mm) e=0.8mm",  peso_unit:  1.500, unidad:"mm"},
  {codigo:"TESR-1/2x0.9",    descripcion:"Tubo ES Redondo 1/2\" (12.70mm) e=0.9mm",  peso_unit:  1.674, unidad:"mm"},
  {codigo:"TESR-1/2x1.0",    descripcion:"Tubo ES Redondo 1/2\" (12.70mm) e=1.0mm",  peso_unit:  1.842, unidad:"mm"},
  {codigo:"TESR-1/2x1.2",    descripcion:"Tubo ES Redondo 1/2\" (12.70mm) e=1.2mm",  peso_unit:  2.196, unidad:"mm"},
  {codigo:"TESR-1/2x1.5",    descripcion:"Tubo ES Redondo 1/2\" (12.70mm) e=1.5mm",  peso_unit:  2.712, unidad:"mm"},
  {codigo:"TESR-1/2x2.0",    descripcion:"Tubo ES Redondo 1/2\" (12.70mm) e=2.0mm",  peso_unit:  3.540, unidad:"mm"},
  {codigo:"TESR-5/8x0.8",    descripcion:"Tubo ES Redondo 5/8\" (15.88mm) e=0.8mm",  peso_unit:  1.860, unidad:"mm"},
  {codigo:"TESR-5/8x0.9",    descripcion:"Tubo ES Redondo 5/8\" (15.88mm) e=0.9mm",  peso_unit:  2.090, unidad:"mm"},
  {codigo:"TESR-5/8x1.0",    descripcion:"Tubo ES Redondo 5/8\" (15.88mm) e=1.0mm",  peso_unit:  2.320, unidad:"mm"},
  {codigo:"TESR-5/8x1.2",    descripcion:"Tubo ES Redondo 5/8\" (15.88mm) e=1.2mm",  peso_unit:  2.760, unidad:"mm"},
  {codigo:"TESR-5/8x1.5",    descripcion:"Tubo ES Redondo 5/8\" (15.88mm) e=1.5mm",  peso_unit:  3.390, unidad:"mm"},
  {codigo:"TESR-5/8x2.0",    descripcion:"Tubo ES Redondo 5/8\" (15.88mm) e=2.0mm",  peso_unit:  4.430, unidad:"mm"},
  {codigo:"TESR-3/4x0.8",    descripcion:"Tubo ES Redondo 3/4\" (19.05mm) e=0.8mm",  peso_unit:  2.240, unidad:"mm"},
  {codigo:"TESR-3/4x0.9",    descripcion:"Tubo ES Redondo 3/4\" (19.05mm) e=0.9mm",  peso_unit:  2.510, unidad:"mm"},
  {codigo:"TESR-3/4x1.0",    descripcion:"Tubo ES Redondo 3/4\" (19.05mm) e=1.0mm",  peso_unit:  2.770, unidad:"mm"},
  {codigo:"TESR-3/4x1.2",    descripcion:"Tubo ES Redondo 3/4\" (19.05mm) e=1.2mm",  peso_unit:  3.320, unidad:"mm"},
  {codigo:"TESR-3/4x1.5",    descripcion:"Tubo ES Redondo 3/4\" (19.05mm) e=1.5mm",  peso_unit:  4.070, unidad:"mm"},
  {codigo:"TESR-3/4x2.0",    descripcion:"Tubo ES Redondo 3/4\" (19.05mm) e=2.0mm",  peso_unit:  5.420, unidad:"mm"},
  {codigo:"TESR-7/8x0.8",    descripcion:"Tubo ES Redondo 7/8\" (22.23mm) e=0.8mm",  peso_unit:  2.610, unidad:"mm"},
  {codigo:"TESR-7/8x0.9",    descripcion:"Tubo ES Redondo 7/8\" (22.23mm) e=0.9mm",  peso_unit:  2.929, unidad:"mm"},
  {codigo:"TESR-7/8x1.0",    descripcion:"Tubo ES Redondo 7/8\" (22.23mm) e=1.0mm",  peso_unit:  3.242, unidad:"mm"},
  {codigo:"TESR-7/8x1.2",    descripcion:"Tubo ES Redondo 7/8\" (22.23mm) e=1.2mm",  peso_unit:  3.876, unidad:"mm"},
  {codigo:"TESR-7/8x1.5",    descripcion:"Tubo ES Redondo 7/8\" (22.23mm) e=1.5mm",  peso_unit:  4.836, unidad:"mm"},
  {codigo:"TESR-7/8x2.0",    descripcion:"Tubo ES Redondo 7/8\" (22.23mm) e=2.0mm",  peso_unit:  6.365, unidad:"mm"},
  {codigo:"TESR-1x0.8",      descripcion:"Tubo ES Redondo 1\" (25.40mm) e=0.8mm",    peso_unit:  3.010, unidad:"mm"},
  {codigo:"TESR-1x0.9",      descripcion:"Tubo ES Redondo 1\" (25.40mm) e=0.9mm",    peso_unit:  3.359, unidad:"mm"},
  {codigo:"TESR-1x1.0",      descripcion:"Tubo ES Redondo 1\" (25.40mm) e=1.0mm",    peso_unit:  3.721, unidad:"mm"},
  {codigo:"TESR-1x1.2",      descripcion:"Tubo ES Redondo 1\" (25.40mm) e=1.2mm",    peso_unit:  4.458, unidad:"mm"},
  {codigo:"TESR-1x1.5",      descripcion:"Tubo ES Redondo 1\" (25.40mm) e=1.5mm",    peso_unit:  5.563, unidad:"mm"},
  {codigo:"TESR-1x2.0",      descripcion:"Tubo ES Redondo 1\" (25.40mm) e=2.0mm",    peso_unit:  7.405, unidad:"mm"},
  {codigo:"TESR-1-1/8x0.9",  descripcion:"Tubo ES Redondo 1 1/8\" (28.58mm) e=0.9mm",peso_unit:  3.779, unidad:"mm"},
  {codigo:"TESR-1-1/8x1.0",  descripcion:"Tubo ES Redondo 1 1/8\" (28.58mm) e=1.0mm",peso_unit:  4.187, unidad:"mm"},
  {codigo:"TESR-1-1/8x1.2",  descripcion:"Tubo ES Redondo 1 1/8\" (28.58mm) e=1.2mm",peso_unit:  5.010, unidad:"mm"},
  {codigo:"TESR-1-1/8x1.5",  descripcion:"Tubo ES Redondo 1 1/8\" (28.58mm) e=1.5mm",peso_unit:  6.245, unidad:"mm"},
  {codigo:"TESR-1-1/8x2.0",  descripcion:"Tubo ES Redondo 1 1/8\" (28.58mm) e=2.0mm",peso_unit:  8.255, unidad:"mm"},
  {codigo:"TESR-1-1/4x0.9",  descripcion:"Tubo ES Redondo 1 1/4\" (31.75mm) e=0.9mm",peso_unit:  4.200, unidad:"mm"},
  {codigo:"TESR-1-1/4x1.0",  descripcion:"Tubo ES Redondo 1 1/4\" (31.75mm) e=1.0mm",peso_unit:  4.662, unidad:"mm"},
  {codigo:"TESR-1-1/4x1.2",  descripcion:"Tubo ES Redondo 1 1/4\" (31.75mm) e=1.2mm",peso_unit:  4.586, unidad:"mm"},
  {codigo:"TESR-1-1/4x1.5",  descripcion:"Tubo ES Redondo 1 1/4\" (31.75mm) e=1.5mm",peso_unit:  6.948, unidad:"mm"},
  {codigo:"TESR-1-1/4x2.0",  descripcion:"Tubo ES Redondo 1 1/4\" (31.75mm) e=2.0mm",peso_unit:  9.210, unidad:"mm"},
  {codigo:"TESR-1-3/8x0.9",  descripcion:"Tubo ES Redondo 1 3/8\" (34.93mm) e=0.9mm",peso_unit:  4.608, unidad:"mm"},
  {codigo:"TESR-1-3/8x1.0",  descripcion:"Tubo ES Redondo 1 3/8\" (34.93mm) e=1.0mm",peso_unit:  5.110, unidad:"mm"},
  {codigo:"TESR-1-3/8x1.2",  descripcion:"Tubo ES Redondo 1 3/8\" (34.93mm) e=1.2mm",peso_unit:  6.132, unidad:"mm"},
  {codigo:"TESR-1-3/8x1.5",  descripcion:"Tubo ES Redondo 1 3/8\" (34.93mm) e=1.5mm",peso_unit:  7.627, unidad:"mm"},
  {codigo:"TESR-1-3/8x2.0",  descripcion:"Tubo ES Redondo 1 3/8\" (34.93mm) e=2.0mm",peso_unit: 10.073, unidad:"mm"},
  {codigo:"TESR-1-1/2x0.9",  descripcion:"Tubo ES Redondo 1 1/2\" (38.10mm) e=0.9mm",peso_unit:  5.070, unidad:"mm"},
  {codigo:"TESR-1-1/2x1.0",  descripcion:"Tubo ES Redondo 1 1/2\" (38.10mm) e=1.0mm",peso_unit:  5.628, unidad:"mm"},
  {codigo:"TESR-1-1/2x1.2",  descripcion:"Tubo ES Redondo 1 1/2\" (38.10mm) e=1.2mm",peso_unit:  6.750, unidad:"mm"},
  {codigo:"TESR-1-1/2x1.5",  descripcion:"Tubo ES Redondo 1 1/2\" (38.10mm) e=1.5mm",peso_unit:  8.400, unidad:"mm"},
  {codigo:"TESR-1-1/2x2.0",  descripcion:"Tubo ES Redondo 1 1/2\" (38.10mm) e=2.0mm",peso_unit: 11.052, unidad:"mm"},
  {codigo:"TESR-1-5/8x0.9",  descripcion:"Tubo ES Redondo 1 5/8\" (41.28mm) e=0.9mm",peso_unit:  5.491, unidad:"mm"},
  {codigo:"TESR-1-5/8x1.0",  descripcion:"Tubo ES Redondo 1 5/8\" (41.28mm) e=1.0mm",peso_unit:  6.083, unidad:"mm"},
  {codigo:"TESR-1-5/8x1.2",  descripcion:"Tubo ES Redondo 1 5/8\" (41.28mm) e=1.2mm",peso_unit:  7.264, unidad:"mm"},
  {codigo:"TESR-1-5/8x1.5",  descripcion:"Tubo ES Redondo 1 5/8\" (41.28mm) e=1.5mm",peso_unit:  9.044, unidad:"mm"},
  {codigo:"TESR-1-5/8x2.0",  descripcion:"Tubo ES Redondo 1 5/8\" (41.28mm) e=2.0mm",peso_unit: 12.047, unidad:"mm"},
  {codigo:"TESR-1-3/4x0.9",  descripcion:"Tubo ES Redondo 1 3/4\" (44.45mm) e=0.9mm",peso_unit:  5.933, unidad:"mm"},
  {codigo:"TESR-1-3/4x1.0",  descripcion:"Tubo ES Redondo 1 3/4\" (44.45mm) e=1.0mm",peso_unit:  6.586, unidad:"mm"},
  {codigo:"TESR-1-3/4x1.2",  descripcion:"Tubo ES Redondo 1 3/4\" (44.45mm) e=1.2mm",peso_unit:  7.896, unidad:"mm"},
  {codigo:"TESR-1-3/4x1.5",  descripcion:"Tubo ES Redondo 1 3/4\" (44.45mm) e=1.5mm",peso_unit:  9.816, unidad:"mm"},
  {codigo:"TESR-1-3/4x2.0",  descripcion:"Tubo ES Redondo 1 3/4\" (44.45mm) e=2.0mm",peso_unit: 13.040, unidad:"mm"},
  {codigo:"TESR-1-7/8x0.9",  descripcion:"Tubo ES Redondo 1 7/8\" (47.63mm) e=0.9mm",peso_unit:  6.374, unidad:"mm"},
  {codigo:"TESR-1-7/8x1.0",  descripcion:"Tubo ES Redondo 1 7/8\" (47.63mm) e=1.0mm",peso_unit:  7.028, unidad:"mm"},
  {codigo:"TESR-1-7/8x1.2",  descripcion:"Tubo ES Redondo 1 7/8\" (47.63mm) e=1.2mm",peso_unit:  8.427, unidad:"mm"},
  {codigo:"TESR-1-7/8x1.5",  descripcion:"Tubo ES Redondo 1 7/8\" (47.63mm) e=1.5mm",peso_unit: 10.525, unidad:"mm"},
  {codigo:"TESR-1-7/8x2.0",  descripcion:"Tubo ES Redondo 1 7/8\" (47.63mm) e=2.0mm",peso_unit: 14.021, unidad:"mm"},
  {codigo:"TESR-2x0.9",      descripcion:"Tubo ES Redondo 2\" (50.80mm) e=0.9mm",    peso_unit:  6.794, unidad:"mm"},
  {codigo:"TESR-2x1.0",      descripcion:"Tubo ES Redondo 2\" (50.80mm) e=1.0mm",    peso_unit:  7.507, unidad:"mm"},
  {codigo:"TESR-2x1.2",      descripcion:"Tubo ES Redondo 2\" (50.80mm) e=1.2mm",    peso_unit:  9.001, unidad:"mm"},
  {codigo:"TESR-2x1.5",      descripcion:"Tubo ES Redondo 2\" (50.80mm) e=1.5mm",    peso_unit: 11.234, unidad:"mm"},
  {codigo:"TESR-2x2.0",      descripcion:"Tubo ES Redondo 2\" (50.80mm) e=2.0mm",    peso_unit: 14.930, unidad:"mm"},
  {codigo:"TESR-2-1/4x1.2",  descripcion:"Tubo ES Redondo 2 1/4\" (57.15mm) e=1.2mm",peso_unit: 10.099, unidad:"mm"},
  {codigo:"TESR-2-1/4x1.5",  descripcion:"Tubo ES Redondo 2 1/4\" (57.15mm) e=1.5mm",peso_unit: 12.642, unidad:"mm"},
  {codigo:"TESR-2-1/4x2.0",  descripcion:"Tubo ES Redondo 2 1/4\" (57.15mm) e=2.0mm",peso_unit: 16.749, unidad:"mm"},
  {codigo:"TESR-2-1/2x1.2",  descripcion:"Tubo ES Redondo 2 1/2\" (63.50mm) e=1.2mm",peso_unit: 11.255, unidad:"mm"},
  {codigo:"TESR-2-1/2x1.5",  descripcion:"Tubo ES Redondo 2 1/2\" (63.50mm) e=1.5mm",peso_unit: 14.033, unidad:"mm"},
  {codigo:"TESR-2-1/2x2.0",  descripcion:"Tubo ES Redondo 2 1/2\" (63.50mm) e=2.0mm",peso_unit: 18.639, unidad:"mm"},
  {codigo:"TESR-3x1.2",      descripcion:"Tubo ES Redondo 3\" (76.20mm) e=1.2mm",    peso_unit: 13.500, unidad:"mm"},
  {codigo:"TESR-3x1.5",      descripcion:"Tubo ES Redondo 3\" (76.20mm) e=1.5mm",    peso_unit: 16.860, unidad:"mm"},
  {codigo:"TESR-3x2.0",      descripcion:"Tubo ES Redondo 3\" (76.20mm) e=2.0mm",    peso_unit: 22.410, unidad:"mm"},
],

"TUBOS ELECTRO SOLDADOS CUADRADOS (SAE 1010 LAF - COPANT 518 TIPO II)": [
  {codigo:"TESC-1/2x0.8",     descripcion:"Tubo ES Cuadrado 1/2\" (12.70mm) e=0.8mm",    peso_unit:  1.860, unidad:"mm"},
  {codigo:"TESC-1/2x0.9",     descripcion:"Tubo ES Cuadrado 1/2\" (12.70mm) e=0.9mm",    peso_unit:  2.089, unidad:"mm"},
  {codigo:"TESC-1/2x1.0",     descripcion:"Tubo ES Cuadrado 1/2\" (12.70mm) e=1.0mm",    peso_unit:  2.315, unidad:"mm"},
  {codigo:"TESC-1/2x1.2",     descripcion:"Tubo ES Cuadrado 1/2\" (12.70mm) e=1.2mm",    peso_unit:  2.756, unidad:"mm"},
  {codigo:"TESC-1/2x1.5",     descripcion:"Tubo ES Cuadrado 1/2\" (12.70mm) e=1.5mm",    peso_unit:  3.392, unidad:"mm"},
  {codigo:"TESC-1/2x2.0",     descripcion:"Tubo ES Cuadrado 1/2\" (12.70mm) e=2.0mm",    peso_unit:  4.426, unidad:"mm"},
  {codigo:"TESC-5/8x0.8",     descripcion:"Tubo ES Cuadrado 5/8\" (15.88mm) e=0.8mm",    peso_unit:  2.240, unidad:"mm"},
  {codigo:"TESC-5/8x0.9",     descripcion:"Tubo ES Cuadrado 5/8\" (15.88mm) e=0.9mm",    peso_unit:  2.510, unidad:"mm"},
  {codigo:"TESC-5/8x1.0",     descripcion:"Tubo ES Cuadrado 5/8\" (15.88mm) e=1.0mm",    peso_unit:  2.770, unidad:"mm"},
  {codigo:"TESC-5/8x1.2",     descripcion:"Tubo ES Cuadrado 5/8\" (15.88mm) e=1.2mm",    peso_unit:  3.320, unidad:"mm"},
  {codigo:"TESC-5/8x1.5",     descripcion:"Tubo ES Cuadrado 5/8\" (15.88mm) e=1.5mm",    peso_unit:  4.070, unidad:"mm"},
  {codigo:"TESC-5/8x2.0",     descripcion:"Tubo ES Cuadrado 5/8\" (15.88mm) e=2.0mm",    peso_unit:  5.420, unidad:"mm"},
  {codigo:"TESC-3/4x0.8",     descripcion:"Tubo ES Cuadrado 3/4\" (19.05mm) e=0.8mm",    peso_unit:  3.000, unidad:"mm"},
  {codigo:"TESC-3/4x0.9",     descripcion:"Tubo ES Cuadrado 3/4\" (19.05mm) e=0.9mm",    peso_unit:  3.200, unidad:"mm"},
  {codigo:"TESC-3/4x1.0",     descripcion:"Tubo ES Cuadrado 3/4\" (19.05mm) e=1.0mm",    peso_unit:  3.481, unidad:"mm"},
  {codigo:"TESC-3/4x1.2",     descripcion:"Tubo ES Cuadrado 3/4\" (19.05mm) e=1.2mm",    peso_unit:  4.270, unidad:"mm"},
  {codigo:"TESC-3/4x1.5",     descripcion:"Tubo ES Cuadrado 3/4\" (19.05mm) e=1.5mm",    peso_unit:  5.330, unidad:"mm"},
  {codigo:"TESC-3/4x2.0",     descripcion:"Tubo ES Cuadrado 3/4\" (19.05mm) e=2.0mm",    peso_unit:  7.310, unidad:"mm"},
  {codigo:"TESC-7/8x0.8",     descripcion:"Tubo ES Cuadrado 7/8\" (22.23mm) e=0.8mm",    peso_unit:  3.390, unidad:"mm"},
  {codigo:"TESC-7/8x0.9",     descripcion:"Tubo ES Cuadrado 7/8\" (22.23mm) e=0.9mm",    peso_unit:  3.779, unidad:"mm"},
  {codigo:"TESC-7/8x1.0",     descripcion:"Tubo ES Cuadrado 7/8\" (22.23mm) e=1.0mm",    peso_unit:  4.187, unidad:"mm"},
  {codigo:"TESC-7/8x1.2",     descripcion:"Tubo ES Cuadrado 7/8\" (22.23mm) e=1.2mm",    peso_unit:  5.010, unidad:"mm"},
  {codigo:"TESC-7/8x1.5",     descripcion:"Tubo ES Cuadrado 7/8\" (22.23mm) e=1.5mm",    peso_unit:  6.245, unidad:"mm"},
  {codigo:"TESC-7/8x2.0",     descripcion:"Tubo ES Cuadrado 7/8\" (22.23mm) e=2.0mm",    peso_unit:  8.255, unidad:"mm"},
  {codigo:"TESC-1x0.8",       descripcion:"Tubo ES Cuadrado 1\" (25.40mm) e=0.8mm",      peso_unit:  3.720, unidad:"mm"},
  {codigo:"TESC-1x0.9",       descripcion:"Tubo ES Cuadrado 1\" (25.40mm) e=0.9mm",      peso_unit:  4.199, unidad:"mm"},
  {codigo:"TESC-1x1.0",       descripcion:"Tubo ES Cuadrado 1\" (25.40mm) e=1.0mm",      peso_unit:  4.660, unidad:"mm"},
  {codigo:"TESC-1x1.2",       descripcion:"Tubo ES Cuadrado 1\" (25.40mm) e=1.2mm",      peso_unit:  5.585, unidad:"mm"},
  {codigo:"TESC-1x1.5",       descripcion:"Tubo ES Cuadrado 1\" (25.40mm) e=1.5mm",      peso_unit:  6.945, unidad:"mm"},
  {codigo:"TESC-1x2.0",       descripcion:"Tubo ES Cuadrado 1\" (25.40mm) e=2.0mm",      peso_unit:  9.212, unidad:"mm"},
  {codigo:"TESC-1-3/16x0.9",  descripcion:"Tubo ES Cuadrado 1 3/16\" (30.16mm) e=0.9mm", peso_unit:  5.071, unidad:"mm"},
  {codigo:"TESC-1-3/16x1.0",  descripcion:"Tubo ES Cuadrado 1 3/16\" (30.16mm) e=1.0mm", peso_unit:  5.629, unidad:"mm"},
  {codigo:"TESC-1-3/16x1.2",  descripcion:"Tubo ES Cuadrado 1 3/16\" (30.16mm) e=1.2mm", peso_unit:  6.747, unidad:"mm"},
  {codigo:"TESC-1-3/16x1.5",  descripcion:"Tubo ES Cuadrado 1 3/16\" (30.16mm) e=1.5mm", peso_unit:  8.398, unidad:"mm"},
  {codigo:"TESC-1-3/16x2.0",  descripcion:"Tubo ES Cuadrado 1 3/16\" (30.16mm) e=2.0mm", peso_unit: 11.054, unidad:"mm"},
  {codigo:"TESC-1-1/4x0.9",   descripcion:"Tubo ES Cuadrado 1 1/4\" (31.75mm) e=0.9mm",  peso_unit:  5.491, unidad:"mm"},
  {codigo:"TESC-1-1/4x1.0",   descripcion:"Tubo ES Cuadrado 1 1/4\" (31.75mm) e=1.0mm",  peso_unit:  6.083, unidad:"mm"},
  {codigo:"TESC-1-1/4x1.2",   descripcion:"Tubo ES Cuadrado 1 1/4\" (31.75mm) e=1.2mm",  peso_unit:  7.264, unidad:"mm"},
  {codigo:"TESC-1-1/4x1.5",   descripcion:"Tubo ES Cuadrado 1 1/4\" (31.75mm) e=1.5mm",  peso_unit:  9.044, unidad:"mm"},
  {codigo:"TESC-1-1/4x2.0",   descripcion:"Tubo ES Cuadrado 1 1/4\" (31.75mm) e=2.0mm",  peso_unit: 12.047, unidad:"mm"},
  {codigo:"TESC-1-1/2x0.9",   descripcion:"Tubo ES Cuadrado 1 1/2\" (38.10mm) e=0.9mm",  peso_unit:  6.439, unidad:"mm"},
  {codigo:"TESC-1-1/2x1.0",   descripcion:"Tubo ES Cuadrado 1 1/2\" (38.10mm) e=1.0mm",  peso_unit:  7.136, unidad:"mm"},
  {codigo:"TESC-1-1/2x1.2",   descripcion:"Tubo ES Cuadrado 1 1/2\" (38.10mm) e=1.2mm",  peso_unit:  8.542, unidad:"mm"},
  {codigo:"TESC-1-1/2x1.5",   descripcion:"Tubo ES Cuadrado 1 1/2\" (38.10mm) e=1.5mm",  peso_unit: 10.659, unidad:"mm"},
  {codigo:"TESC-1-1/2x2.0",   descripcion:"Tubo ES Cuadrado 1 1/2\" (38.10mm) e=2.0mm",  peso_unit: 14.021, unidad:"mm"},
  {codigo:"TESC-2x1.2",       descripcion:"Tubo ES Cuadrado 2\" (50.80mm) e=1.2mm",      peso_unit: 11.255, unidad:"mm"},
  {codigo:"TESC-2x1.5",       descripcion:"Tubo ES Cuadrado 2\" (50.80mm) e=1.5mm",      peso_unit: 14.033, unidad:"mm"},
  {codigo:"TESC-2x2.0",       descripcion:"Tubo ES Cuadrado 2\" (50.80mm) e=2.0mm",      peso_unit: 18.639, unidad:"mm"},
  {codigo:"TESC-15mmx0.8",    descripcion:"Tubo ES Cuadrado 15mm e=0.8mm",               peso_unit:  2.240, unidad:"mm"},
  {codigo:"TESC-15mmx0.9",    descripcion:"Tubo ES Cuadrado 15mm e=0.9mm",               peso_unit:  2.510, unidad:"mm"},
  {codigo:"TESC-15mmx1.0",    descripcion:"Tubo ES Cuadrado 15mm e=1.0mm",               peso_unit:  2.770, unidad:"mm"},
  {codigo:"TESC-15mmx1.2",    descripcion:"Tubo ES Cuadrado 15mm e=1.2mm",               peso_unit:  3.320, unidad:"mm"},
  {codigo:"TESC-15mmx1.5",    descripcion:"Tubo ES Cuadrado 15mm e=1.5mm",               peso_unit:  4.070, unidad:"mm"},
  {codigo:"TESC-20mmx0.8",    descripcion:"Tubo ES Cuadrado 20mm e=0.8mm",               peso_unit:  3.010, unidad:"mm"},
  {codigo:"TESC-20mmx0.9",    descripcion:"Tubo ES Cuadrado 20mm e=0.9mm",               peso_unit:  3.359, unidad:"mm"},
  {codigo:"TESC-20mmx1.0",    descripcion:"Tubo ES Cuadrado 20mm e=1.0mm",               peso_unit:  3.721, unidad:"mm"},
  {codigo:"TESC-20mmx1.2",    descripcion:"Tubo ES Cuadrado 20mm e=1.2mm",               peso_unit:  4.458, unidad:"mm"},
  {codigo:"TESC-20mmx1.5",    descripcion:"Tubo ES Cuadrado 20mm e=1.5mm",               peso_unit:  5.563, unidad:"mm"},
  {codigo:"TESC-20mmx2.0",    descripcion:"Tubo ES Cuadrado 20mm e=2.0mm",               peso_unit:  7.405, unidad:"mm"},
  {codigo:"TESC-25mmx0.8",    descripcion:"Tubo ES Cuadrado 25mm e=0.8mm",               peso_unit:  3.720, unidad:"mm"},
  {codigo:"TESC-25mmx0.9",    descripcion:"Tubo ES Cuadrado 25mm e=0.9mm",               peso_unit:  4.199, unidad:"mm"},
  {codigo:"TESC-25mmx1.0",    descripcion:"Tubo ES Cuadrado 25mm e=1.0mm",               peso_unit:  4.660, unidad:"mm"},
  {codigo:"TESC-25mmx1.2",    descripcion:"Tubo ES Cuadrado 25mm e=1.2mm",               peso_unit:  5.585, unidad:"mm"},
  {codigo:"TESC-25mmx1.5",    descripcion:"Tubo ES Cuadrado 25mm e=1.5mm",               peso_unit:  6.945, unidad:"mm"},
  {codigo:"TESC-25mmx2.0",    descripcion:"Tubo ES Cuadrado 25mm e=2.0mm",               peso_unit:  9.212, unidad:"mm"},
  {codigo:"TESC-30mmx0.9",    descripcion:"Tubo ES Cuadrado 30mm e=0.9mm",               peso_unit:  5.071, unidad:"mm"},
  {codigo:"TESC-30mmx1.0",    descripcion:"Tubo ES Cuadrado 30mm e=1.0mm",               peso_unit:  5.629, unidad:"mm"},
  {codigo:"TESC-30mmx1.2",    descripcion:"Tubo ES Cuadrado 30mm e=1.2mm",               peso_unit:  6.747, unidad:"mm"},
  {codigo:"TESC-30mmx1.5",    descripcion:"Tubo ES Cuadrado 30mm e=1.5mm",               peso_unit:  8.398, unidad:"mm"},
  {codigo:"TESC-30mmx2.0",    descripcion:"Tubo ES Cuadrado 30mm e=2.0mm",               peso_unit: 11.054, unidad:"mm"},
  {codigo:"TESC-40mmx0.9",    descripcion:"Tubo ES Cuadrado 40mm e=0.9mm",               peso_unit:  6.794, unidad:"mm"},
  {codigo:"TESC-40mmx1.0",    descripcion:"Tubo ES Cuadrado 40mm e=1.0mm",               peso_unit:  7.507, unidad:"mm"},
  {codigo:"TESC-40mmx1.2",    descripcion:"Tubo ES Cuadrado 40mm e=1.2mm",               peso_unit:  9.001, unidad:"mm"},
  {codigo:"TESC-40mmx1.5",    descripcion:"Tubo ES Cuadrado 40mm e=1.5mm",               peso_unit: 11.234, unidad:"mm"},
  {codigo:"TESC-40mmx2.0",    descripcion:"Tubo ES Cuadrado 40mm e=2.0mm",               peso_unit: 14.930, unidad:"mm"},
  {codigo:"TESC-50mmx1.2",    descripcion:"Tubo ES Cuadrado 50mm e=1.2mm",               peso_unit: 11.255, unidad:"mm"},
  {codigo:"TESC-50mmx1.5",    descripcion:"Tubo ES Cuadrado 50mm e=1.5mm",               peso_unit: 14.033, unidad:"mm"},
  {codigo:"TESC-50mmx2.0",    descripcion:"Tubo ES Cuadrado 50mm e=2.0mm",               peso_unit: 18.639, unidad:"mm"},
  {codigo:"TESC-60mmx1.2",    descripcion:"Tubo ES Cuadrado 60mm e=1.2mm",               peso_unit: 13.495, unidad:"mm"},
  {codigo:"TESC-60mmx1.5",    descripcion:"Tubo ES Cuadrado 60mm e=1.5mm",               peso_unit: 16.859, unidad:"mm"},
  {codigo:"TESC-60mmx2.0",    descripcion:"Tubo ES Cuadrado 60mm e=2.0mm",               peso_unit: 22.407, unidad:"mm"},
],

"TUBOS ELECTRO SOLDADOS RECTANGULARES (SAE 1010 LAF - COPANT 518 TIPO II)": [
  {codigo:"TESR-1/2x3/4x0.9",  descripcion:"Tubo ES Rectangular 1/2\" × 3/4\" (12.70×19.05mm) e=0.9mm",  peso_unit: 2.692, unidad:"mm"},
  {codigo:"TESR-1/2x1x0.8",    descripcion:"Tubo ES Rectangular 1/2\" × 1\" (12.70×25.40mm) e=0.8mm",    peso_unit: 3.010, unidad:"mm"},
  {codigo:"TESR-1/2x1x0.9",    descripcion:"Tubo ES Rectangular 1/2\" × 1\" (12.70×25.40mm) e=0.9mm",    peso_unit: 3.359, unidad:"mm"},
  {codigo:"TESR-1/2x1x1.0",    descripcion:"Tubo ES Rectangular 1/2\" × 1\" (12.70×25.40mm) e=1.0mm",    peso_unit: 3.721, unidad:"mm"},
  {codigo:"TESR-1/2x1x1.2",    descripcion:"Tubo ES Rectangular 1/2\" × 1\" (12.70×25.40mm) e=1.2mm",    peso_unit: 4.458, unidad:"mm"},
  {codigo:"TESR-1/2x1x1.5",    descripcion:"Tubo ES Rectangular 1/2\" × 1\" (12.70×25.40mm) e=1.5mm",    peso_unit: 5.563, unidad:"mm"},
  {codigo:"TESR-1/2x1x2.0",    descripcion:"Tubo ES Rectangular 1/2\" × 1\" (12.70×25.40mm) e=2.0mm",    peso_unit: 7.405, unidad:"mm"},
  {codigo:"TESR-1/2x1-1/2x0.9",descripcion:"Tubo ES Rectangular 1/2\" × 1 1/2\" (12.70×38.10mm) e=0.9mm",peso_unit: 4.199, unidad:"mm"},
  {codigo:"TESR-1/2x1-1/2x1.0",descripcion:"Tubo ES Rectangular 1/2\" × 1 1/2\" (12.70×38.10mm) e=1.0mm",peso_unit: 4.660, unidad:"mm"},
  {codigo:"TESR-1/2x1-1/2x1.2",descripcion:"Tubo ES Rectangular 1/2\" × 1 1/2\" (12.70×38.10mm) e=1.2mm",peso_unit: 5.585, unidad:"mm"},
  {codigo:"TESR-1/2x1-1/2x1.5",descripcion:"Tubo ES Rectangular 1/2\" × 1 1/2\" (12.70×38.10mm) e=1.5mm",peso_unit: 6.945, unidad:"mm"},
  {codigo:"TESR-1/2x1-1/2x2.0",descripcion:"Tubo ES Rectangular 1/2\" × 1 1/2\" (12.70×38.10mm) e=2.0mm",peso_unit: 9.212, unidad:"mm"},
  {codigo:"TESR-1x2x0.9",      descripcion:"Tubo ES Rectangular 1\" × 2\" (25.40×50.80mm) e=0.9mm",      peso_unit:  6.438, unidad:"mm"},
  {codigo:"TESR-1x2x1.0",      descripcion:"Tubo ES Rectangular 1\" × 2\" (25.40×50.80mm) e=1.0mm",      peso_unit:  7.136, unidad:"mm"},
  {codigo:"TESR-1x2x1.2",      descripcion:"Tubo ES Rectangular 1\" × 2\" (25.40×50.80mm) e=1.2mm",      peso_unit:  8.541, unidad:"mm"},
  {codigo:"TESR-1x2x1.5",      descripcion:"Tubo ES Rectangular 1\" × 2\" (25.40×50.80mm) e=1.5mm",      peso_unit: 10.659, unidad:"mm"},
  {codigo:"TESR-1x2x2.0",      descripcion:"Tubo ES Rectangular 1\" × 2\" (25.40×50.80mm) e=2.0mm",      peso_unit: 14.188, unidad:"mm"},
  {codigo:"TESR-20x40x0.9",    descripcion:"Tubo ES Rectangular 20×40mm e=0.9mm",                         peso_unit:  5.071, unidad:"mm"},
  {codigo:"TESR-20x40x1.0",    descripcion:"Tubo ES Rectangular 20×40mm e=1.0mm",                         peso_unit:  5.629, unidad:"mm"},
  {codigo:"TESR-20x40x1.2",    descripcion:"Tubo ES Rectangular 20×40mm e=1.2mm",                         peso_unit:  6.747, unidad:"mm"},
  {codigo:"TESR-20x40x1.5",    descripcion:"Tubo ES Rectangular 20×40mm e=1.5mm",                         peso_unit:  8.398, unidad:"mm"},
  {codigo:"TESR-20x40x2.0",    descripcion:"Tubo ES Rectangular 20×40mm e=2.0mm",                         peso_unit: 11.054, unidad:"mm"},
  {codigo:"TESR-20x50x0.9",    descripcion:"Tubo ES Rectangular 20×50mm e=0.9mm",                         peso_unit:  5.930, unidad:"mm"},
  {codigo:"TESR-20x50x1.0",    descripcion:"Tubo ES Rectangular 20×50mm e=1.0mm",                         peso_unit:  6.586, unidad:"mm"},
  {codigo:"TESR-20x50x1.2",    descripcion:"Tubo ES Rectangular 20×50mm e=1.2mm",                         peso_unit:  7.896, unidad:"mm"},
  {codigo:"TESR-20x50x1.5",    descripcion:"Tubo ES Rectangular 20×50mm e=1.5mm",                         peso_unit:  9.816, unidad:"mm"},
  {codigo:"TESR-20x50x2.0",    descripcion:"Tubo ES Rectangular 20×50mm e=2.0mm",                         peso_unit: 13.040, unidad:"mm"},
  {codigo:"TESR-30x40x0.9",    descripcion:"Tubo ES Rectangular 30×40mm e=0.9mm",                         peso_unit:  5.930, unidad:"mm"},
  {codigo:"TESR-30x40x1.0",    descripcion:"Tubo ES Rectangular 30×40mm e=1.0mm",                         peso_unit:  6.586, unidad:"mm"},
  {codigo:"TESR-30x40x1.2",    descripcion:"Tubo ES Rectangular 30×40mm e=1.2mm",                         peso_unit:  7.896, unidad:"mm"},
  {codigo:"TESR-30x40x1.5",    descripcion:"Tubo ES Rectangular 30×40mm e=1.5mm",                         peso_unit:  9.916, unidad:"mm"},
  {codigo:"TESR-30x40x2.0",    descripcion:"Tubo ES Rectangular 30×40mm e=2.0mm",                         peso_unit: 13.040, unidad:"mm"},
  {codigo:"TESR-30x50x0.9",    descripcion:"Tubo ES Rectangular 30×50mm e=0.9mm",                         peso_unit:  6.794, unidad:"mm"},
  {codigo:"TESR-30x50x1.0",    descripcion:"Tubo ES Rectangular 30×50mm e=1.0mm",                         peso_unit:  7.507, unidad:"mm"},
  {codigo:"TESR-30x50x1.2",    descripcion:"Tubo ES Rectangular 30×50mm e=1.2mm",                         peso_unit:  9.001, unidad:"mm"},
  {codigo:"TESR-30x50x1.5",    descripcion:"Tubo ES Rectangular 30×50mm e=1.5mm",                         peso_unit: 11.234, unidad:"mm"},
  {codigo:"TESR-30x50x2.0",    descripcion:"Tubo ES Rectangular 30×50mm e=2.0mm",                         peso_unit: 14.930, unidad:"mm"},
  {codigo:"TESR-40x50x1.2",    descripcion:"Tubo ES Rectangular 40×50mm e=1.2mm",                         peso_unit: 10.080, unidad:"mm"},
  {codigo:"TESR-40x50x1.5",    descripcion:"Tubo ES Rectangular 40×50mm e=1.5mm",                         peso_unit: 12.660, unidad:"mm"},
  {codigo:"TESR-40x50x2.0",    descripcion:"Tubo ES Rectangular 40×50mm e=2.0mm",                         peso_unit: 16.749, unidad:"mm"},
  {codigo:"TESR-40x60x1.2",    descripcion:"Tubo ES Rectangular 40×60mm e=1.2mm",                         peso_unit: 11.255, unidad:"mm"},
  {codigo:"TESR-40x60x1.5",    descripcion:"Tubo ES Rectangular 40×60mm e=1.5mm",                         peso_unit: 14.033, unidad:"mm"},
  {codigo:"TESR-40x60x2.0",    descripcion:"Tubo ES Rectangular 40×60mm e=2.0mm",                         peso_unit: 18.639, unidad:"mm"},
  {codigo:"TESR-40x80x1.5",    descripcion:"Tubo ES Rectangular 40×80mm e=1.5mm",                         peso_unit: 16.860, unidad:"mm"},
  {codigo:"TESR-40x80x2.0",    descripcion:"Tubo ES Rectangular 40×80mm e=2.0mm",                         peso_unit: 22.410, unidad:"mm"},
  {codigo:"TESR-50x70x1.5",    descripcion:"Tubo ES Rectangular 50×70mm e=1.5mm",                         peso_unit: 16.860, unidad:"mm"},
  {codigo:"TESR-50x70x2.0",    descripcion:"Tubo ES Rectangular 50×70mm e=2.0mm",                         peso_unit: 22.410, unidad:"mm"},
  {codigo:"TESR-50x100x2.0",   descripcion:"Tubo ES Rectangular 50×100mm e=2.0mm",                        peso_unit: 28.080, unidad:"mm"},
],
};

// ─── ESTADO ───
let items = [];
let activeId = null;
let editingRate = null;
let toastTm = null;

// Hidden rate inputs
let puFab = 15.20, puMont = 5.20;

// ─── DOM ───
const $cat = document.getElementById('categoria');
const $prod = document.getElementById('producto');
const $search = document.getElementById('search');
const $preview = document.getElementById('preview');
const $cant = document.getElementById('cantidad');
const $unidad = document.getElementById('unidad');
const $est = document.getElementById('est_box');
const $estV = document.getElementById('est_value');

// Populate categories
Object.keys(CAT).forEach(k => {
  const o = document.createElement('option');
  o.value = k; o.textContent = k;
  $cat.appendChild(o);
});

// ─── WORKFLOW ───
let currentStep = 0;
function goStep(n) {
  currentStep = n;
  document.querySelectorAll('.step-panel').forEach((p,i) => p.classList.toggle('active', i===n));
  document.querySelectorAll('.wf-step').forEach((s,i) => {
    s.classList.toggle('active', i===n);
    s.classList.toggle('done', i<n);
  });
  document.querySelectorAll('.wf-connector').forEach((c,i) => c.classList.toggle('done', i<n));
}

// ─── ITEMS ───
function openItemModal() {
  const modal = document.getElementById('item_modal');
  const inp = document.getElementById('modal_input');
  inp.value = document.getElementById('new_item_name').value;
  modal.classList.add('open');
  setTimeout(() => inp.focus(), 120);
}
function closeItemModal() {
  document.getElementById('item_modal').classList.remove('open');
}
function createItem() {
  const name = document.getElementById('modal_input').value.trim();
  if (!name) { toast('Escribe un nombre para el ítem', 1); return; }
  const id = 'i' + Date.now();
  items.push({ id, nombre: name, subs: [] });
  setActive(id);
  document.getElementById('new_item_name').value = '';
  closeItemModal();
  renderItems(); renderTable(); updateStats();
  toast('Ítem creado');
  goStep(1);
}
function deleteItem(id) {
  const it = items.find(x => x.id === id);
  if (!it) return;
  if (it.subs.length > 0 && !confirm(`¿Eliminar "${it.nombre}" con sus ${it.subs.length} sub-ítem(s)?`)) return;
  items = items.filter(x => x.id !== id);
  if (activeId === id) setActive(items.length ? items[items.length-1].id : null);
  else setActive(activeId);
  renderItems(); renderTable(); updateStats();
  toast('Ítem eliminado');
}
function setActive(id) {
  activeId = id;
  const it = items.find(x => x.id === id);
  const banner = document.getElementById('active_banner');
  if (it) {
    const idx = items.indexOf(it);
    banner.className = 'active-banner';
    banner.innerHTML = `<span class="ab-name">${idx+1}. ${it.nombre}</span><span class="ab-count">${it.subs.length} sub-ítem(s)</span>`;
  } else {
    banner.className = 'active-banner none';
    banner.innerHTML = '<span class="ab-name">Selecciona un ítem en el paso 1</span>';
  }
  renderItems();
}
function renderItems() {
  const list = document.getElementById('items_list');
  const badge = document.getElementById('item_count_badge');
  badge.textContent = items.length ? `(${items.length})` : '';
  if (!items.length) {
    list.innerHTML = `<div class="empty"><div class="empty-icon">📦</div><p>Aún no tienes ítems.<br><b>Crea el primero</b> arriba.</p></div>`;
    return;
  }
  list.innerHTML = '';
  items.forEach((it, idx) => {
    const total = it.subs.reduce((s,x) => s+x.sub, 0);
    const card = document.createElement('div');
    card.className = 'item-card' + (it.id===activeId ? ' active' : '');
    card.innerHTML = `
      <div class="item-card-header" onclick="setActive('${it.id}');goStep(1)">
        <div class="item-num">${idx+1}</div>
        <div class="item-info">
          <div class="item-name">${it.nombre}</div>
          <div class="item-meta">${it.subs.length} sub-ítem(s)${it.subs.length?` · ${it.subs.reduce((s,x)=>s+x.pesoT,0).toFixed(2)} kg`:''}</div>
        </div>
        ${total>0?`<div class="item-total">S/ ${total.toFixed(2)}</div>`:''}
        <button class="item-del" onclick="event.stopPropagation();deleteItem('${it.id}')" title="Eliminar ítem">✕</button>
      </div>
      <div class="item-footer">
        <button class="item-action" onclick="setActive('${it.id}');goStep(1)">+ Agregar sub-ítem</button>
      </div>
    `;
    list.appendChild(card);
  });
}

// ─── CATÁLOGO ───
let srchTm = null;
function debounceSearch() {
  clearTimeout(srchTm);
  const term = $search.value.trim().toLowerCase();
  srchTm = setTimeout(() => {
    if (term.length < 2) { onCategoryChange(); return; }
    let res = [];
    Object.entries(CAT).forEach(([cat,prods]) => prods.forEach(p => {
      if (p.descripcion.toLowerCase().includes(term)||p.codigo.toLowerCase().includes(term))
        res.push({...p, cat});
    }));
    $prod.innerHTML = '<option value="">— Resultados —</option>';
    res.slice(0,20).forEach(p => {
      const o = document.createElement('option');
      o.value = p.codigo;
      o.setAttribute('data-cat', p.cat);
      o.textContent = `${p.descripcion}`;
      $prod.appendChild(o);
    });
    updatePreview(null);
  }, 280);
}
function onCategoryChange() {
  const cat = $cat.value;
  $prod.innerHTML = '<option value="">— Seleccionar producto —</option>';
  if (!cat) { updatePreview(null); return; }
  CAT[cat].forEach(p => {
    const o = document.createElement('option');
    o.value = p.codigo; o.textContent = p.descripcion;
    $prod.appendChild(o);
  });
  updatePreview(null);
}
function onProductChange() {
  const p = getProduct();
  updatePreview(p);
  if (p) { $cant.focus(); updateEstimated(); }
}
function getProduct() {
  const cod = $prod.value;
  const opt = $prod.options[$prod.selectedIndex];
  if (!cod||!opt) return null;
  const cat = opt.getAttribute('data-cat') || $cat.value;
  if (!cat) return null;
  return CAT[cat]?.find(p => p.codigo===cod);
}
function updatePreview(p) {
  if (!p) {
    $preview.className = 'product-card';
    $preview.innerHTML = '<span class="product-card-text">Selecciona un producto para ver sus detalles</span>';
    $unidad.value = '';
    $est.style.display = 'none';
    return;
  }
  $preview.className = 'product-card has-product';
  $preview.innerHTML = `<span class="product-card-text">${p.descripcion}</span><span class="product-badge">${p.peso_unit} kg/${p.unidad}</span>`;
  $unidad.value = p.unidad;
}
function addQty(v) {
  $cant.value = (parseFloat($cant.value)||0)+v;
  updateEstimated();
}
function updateEstimated() {
  const p = getProduct();
  const c = parseFloat($cant.value)||0;
  if (!p||c<=0) { $est.style.display='none'; return; }
  const sub = c*p.peso_unit*(puFab+puMont);
  $estV.textContent = `S/ ${sub.toFixed(2)}`;
  $est.style.display = 'flex';
}
$cant.addEventListener('input', updateEstimated);

// ─── SUB-ÍTEMS ───
function addSubItem() {
  if (!activeId) { toast('Selecciona un ítem primero', 1); return; }
  const p = getProduct();
  if (!p) { toast('Selecciona un producto', 1); return; }
  const cant = parseFloat($cant.value);
  if (!cant||cant<=0) { toast('Ingresa una cantidad válida', 1); return; }
  const pesoT = cant * p.peso_unit;
  const sub = pesoT * (puFab + puMont);
  const item = items.find(x => x.id===activeId);
  item.subs.push({
    id: 's'+Date.now(),
    desc: p.descripcion,
    cant, unidad:p.unidad,
    pesoU:p.peso_unit, pesoT, puFab, puMont, sub
  });
  renderTable(); updateStats(); renderItems();
  setActive(activeId);
  $cant.value = '';
  $est.style.display = 'none';
  toast('Sub-ítem agregado ✓');
  $cant.focus();
}
function deleteSub(itemId, subId) {
  const it = items.find(x => x.id===itemId);
  if (!it) return;
  it.subs = it.subs.filter(s => s.id!==subId);
  renderTable(); updateStats(); renderItems(); setActive(activeId);
  toast('Sub-ítem eliminado');
}
function undoLast() {
  for (let i=items.length-1; i>=0; i--) {
    if (items[i].subs.length) {
      const last = items[i].subs[items[i].subs.length-1];
      deleteSub(items[i].id, last.id);
      return;
    }
  }
  toast('No hay más sub-ítems', 1);
}
function clearAll() {
  if (!items.length) return;
  if (!confirm('¿Eliminar todos los ítems y sub-ítems?')) return;
  items = []; activeId = null;
  setActive(null); renderItems(); renderTable(); updateStats();
  toast('Todo limpiado');
}

// ─── TABLA ───
function renderTable() {
  const tbody = document.getElementById('tbody');
  const empty = document.getElementById('empty_table');
  tbody.innerHTML = '';
  if (!items.length) { empty.style.display='flex'; return; }
  empty.style.display = 'none';
  items.forEach((it,ii) => {
    const iPeso = it.subs.reduce((s,x)=>s+x.pesoT,0);
    const iSub  = it.subs.reduce((s,x)=>s+x.sub,0);
    // main row
    const mr = document.createElement('tr');
    mr.className = 'row-main';
    mr.onclick = () => { setActive(it.id); goStep(1); };
    mr.innerHTML = `
      <td class="row-main-num r">${ii+1}</td>
      <td class="row-main-name">${it.nombre}</td>
      <td class="c" colspan="2" style="color:var(--text-dim);font-size:11px">${it.subs.length} sub-ítem(s)</td>
      <td></td>
      <td class="r" style="font-family:'DM Mono',monospace;color:var(--blue-text);font-size:12px">${iPeso>0?iPeso.toFixed(3):''}</td>
      <td></td><td></td>
      <td class="row-main-total r">${iSub>0?'S/ '+iSub.toFixed(2):''}</td>
      <td class="c"><button class="del" onclick="event.stopPropagation();deleteItem('${it.id}')">✕</button></td>
    `;
    tbody.appendChild(mr);
    it.subs.forEach((s,si) => {
      const sr = document.createElement('tr');
      sr.className = 'row-sub';
      sr.innerHTML = `
        <td class="sub-idx">${ii+1}.${si+1}</td>
        <td class="sub-desc">${s.desc}</td>
        <td class="c">${s.cant}</td>
        <td class="c" style="color:var(--text-dim)">${s.unidad}</td>
        <td class="sub-weight">${s.pesoU.toFixed(3)}</td>
        <td class="sub-weight">${s.pesoT.toFixed(3)}</td>
        <td class="sub-price">${s.puFab.toFixed(2)}</td>
        <td class="sub-price">${s.puMont.toFixed(2)}</td>
        <td class="sub-total">${s.sub.toFixed(2)}</td>
        <td class="c"><button class="del" onclick="deleteSub('${it.id}','${s.id}')">✕</button></td>
      `;
      tbody.appendChild(sr);
    });
  });
}

// ─── STATS ───
function updateStats() {
  const all = items.flatMap(x=>x.subs);
  const kg   = all.reduce((s,x)=>s+x.pesoT,0);
  const fab  = all.reduce((s,x)=>s+x.pesoT*x.puFab,0);
  const mont = all.reduce((s,x)=>s+x.pesoT*x.puMont,0);
  const tot  = all.reduce((s,x)=>s+x.sub,0);
  const ts = 'S/ '+tot.toFixed(2);
  document.getElementById('s_items').textContent = items.length;
  document.getElementById('s_subs').textContent = all.length;
  document.getElementById('s_kg').textContent = kg.toFixed(3);
  document.getElementById('s_fab').textContent = fab.toFixed(2);
  document.getElementById('s_mont').textContent = mont.toFixed(2);
  document.getElementById('s_total').textContent = ts;
  document.getElementById('top_total').textContent = ts;
  document.getElementById('footer_total').textContent = 'TOTAL: '+ts;
  document.getElementById('sum_fab').textContent = 'S/ '+fab.toFixed(2);
  document.getElementById('sum_mont').textContent = 'S/ '+mont.toFixed(2);
  document.getElementById('sum_total').textContent = ts;
}

// ─── TARIFAS ───
function openRateModal(which) {
  editingRate = which;
  const isFab = which==='fab';
  document.getElementById('rate_modal_title').textContent = isFab ? 'Tarifa de Fabricación' : 'Tarifa de Montaje';
  document.getElementById('rate_modal_hint').textContent = isFab
    ? 'Precio por kg de acero fabricado. Se aplicará a todos los sub-ítems.'
    : 'Precio por kg de acero montado. Se aplicará a todos los sub-ítems.';
  const inp = document.getElementById('rate_input');
  inp.value = isFab ? puFab : puMont;
  document.getElementById('rate_modal').classList.add('open');
  setTimeout(() => { inp.focus(); inp.select(); }, 120);
}
function closeRateModal() {
  document.getElementById('rate_modal').classList.remove('open');
  editingRate = null;
}
function saveRate() {
  const v = parseFloat(document.getElementById('rate_input').value)||0;
  if (editingRate==='fab') {
    puFab = v;
    document.getElementById('fab_disp').textContent = v.toFixed(2);
  } else {
    puMont = v;
    document.getElementById('mont_disp').textContent = v.toFixed(2);
  }
  items.forEach(it => it.subs.forEach(s => {
    s.puFab = puFab; s.puMont = puMont;
    s.sub = s.pesoT*(puFab+puMont);
  }));
  renderTable(); updateStats(); renderItems(); updateEstimated();
  closeRateModal();
  toast('Tarifas actualizadas y recalculadas ✓');
}
document.getElementById('rate_input').addEventListener('keydown', e => { if(e.key==='Enter') saveRate(); });

// ─── EXPORT XLSX (reemplaza exportCSV) ───
function exportXLSX() {
  const all = items.flatMap(x => x.subs);
  if (!all.length) { toast('No hay datos para exportar', 1); return; }



  const fecha = new Date().toLocaleDateString('es-PE');
  const notas = document.getElementById('notas').value;

  const wb = XLSX.utils.book_new();
  const ws = {};
  const merges = [];

  // ── Helpers ──
  let R = 0; // fila actual (0-indexed)

  function cell(c, r, v, s) {
    const addr = XLSX.utils.encode_cell({ c, r });
    ws[addr] = { v, t: typeof v === 'number' ? 'n' : 's', s };
  }
  function numCell(c, r, v, fmt, s) {
    const addr = XLSX.utils.encode_cell({ c, r });
    ws[addr] = { v, t: 'n', z: fmt || '#,##0.00', s };
  }
  function fmlaCell(c, r, f, fmt, s) {
    const addr = XLSX.utils.encode_cell({ c, r });
    ws[addr] = { f, t: 'n', z: fmt || '#,##0.00', s };
  }
  function merge(c1, r1, c2, r2) {
    merges.push({ s: { c: c1, r: r1 }, e: { c: c2, r: r2 } });
  }

  // ── Estilos ──
  const COLS = 8; // 0..8 (9 columnas: Item, Desc, Cant, Und, PesoU, PesoT, PU Fab, PU Mont, Subtotal)

  const sBrand = {
    font: { bold: true, sz: 14, color: { rgb: 'FFFFFF' } },
    fill: { fgColor: { rgb: '1A3A5C' } },
    alignment: { horizontal: 'center', vertical: 'center' }
  };
  const sMeta = {
    font: { sz: 10, color: { rgb: '333333' } },
    fill: { fgColor: { rgb: 'EEF2F7' } },
    alignment: { horizontal: 'left' }
  };
  const sMetaLabel = {
    font: { bold: true, sz: 10, color: { rgb: '1A3A5C' } },
    fill: { fgColor: { rgb: 'EEF2F7' } }
  };
  const sHeader = {
    font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } },
    fill: { fgColor: { rgb: '2E6DA4' } },
    alignment: { horizontal: 'center', wrapText: true },
    border: { bottom: { style: 'thin', color: { rgb: 'FFFFFF' } } }
  };
  const sItemGroup = {
    font: { bold: true, sz: 10, color: { rgb: '1A3A5C' } },
    fill: { fgColor: { rgb: 'D6E4F0' } },
    alignment: { horizontal: 'left' }
  };
  const sItemGroupNum = {
    font: { bold: true, sz: 10, color: { rgb: '1A3A5C' } },
    fill: { fgColor: { rgb: 'D6E4F0' } },
    alignment: { horizontal: 'right' }
  };
  const sSub = {
    font: { sz: 9, color: { rgb: '222222' } },
    fill: { fgColor: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'left', indent: 1 }
  };
  const sSubNum = {
    font: { sz: 9, color: { rgb: '222222' } },
    fill: { fgColor: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'right' }
  };
  const sSubFormula = {
    font: { sz: 9, color: { rgb: '000099' } }, // azul = fórmula
    fill: { fgColor: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'right' }
  };
  const sTotal = {
    font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
    fill: { fgColor: { rgb: '1A3A5C' } },
    alignment: { horizontal: 'right' }
  };
  const sTotalLabel = {
    font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
    fill: { fgColor: { rgb: '1A3A5C' } },
    alignment: { horizontal: 'center' }
  };

  // ══════════════════════════════════════════
  // FILA 0 — Título COMASA PRO
  // ══════════════════════════════════════════
  merge(0, R, COLS, R);
  cell(0, R, '🏗  COMASA PRO — Cotización', sBrand);
  for (let c = 1; c <= COLS; c++) cell(c, R, '', sBrand);
  R++;

  // ── Metadatos ──
  const metaRows = [
    ['Fecha', fecha],
    notas ? ['Notas', notas] : null,
    ['Tarifa Fabricación', `S/ ${puFab.toFixed(2)} / kg`],
    ['Tarifa Montaje', `S/ ${puMont.toFixed(2)} / kg`],
  ].filter(Boolean);

  metaRows.forEach(([lbl, val]) => {
    merge(1, R, COLS, R);
    cell(0, R, lbl, sMetaLabel);
    cell(1, R, val, sMeta);
    for (let c = 2; c <= COLS; c++) cell(c, R, '', sMeta);
    R++;
  });

  R++; // espacio

  // ══════════════════════════════════════════
  // ENCABEZADOS
  // ══════════════════════════════════════════
  const headers = ['Item', 'Descripción', 'Cant.', 'Unidad', 'Peso Unit (kg)', 'Peso Total (kg)', 'PU Fab (S/)', 'PU Mont (S/)', 'Subtotal (S/)'];
  headers.forEach((h, c) => cell(c, R, h, sHeader));
  const headerRow = R;
  R++;

  // ══════════════════════════════════════════
  // DATOS + FÓRMULAS
  // ══════════════════════════════════════════
  const itemTotalRows = []; // para fórmula de gran total

  items.forEach((it, i) => {
    const subRows = [];

    // — Subitems —
    it.subs.forEach((s, j) => {
      const exR = R + 1; // Excel row (1-indexed)
      cell(0, R, `${i + 1}.${j + 1}`, sSub);
      cell(1, R, s.desc, sSub);
      numCell(2, R, s.cant, '0.##', sSubNum);
      cell(3, R, s.unidad, sSub);
      numCell(4, R, s.pesoU, '0.000', sSubNum);

      // PesoTotal = cant × pesoUnit (fórmula)
      fmlaCell(5, R, `C${exR}*E${exR}`, '0.000', sSubFormula);

      // PU Fab y Mont como inputs (azul = hardcode editable)
      numCell(6, R, s.puFab, '"S/"#,##0.00', {
        ...sSubNum, font: { sz: 9, color: { rgb: '0000CC' } }
      });
      numCell(7, R, s.puMont, '"S/"#,##0.00', {
        ...sSubNum, font: { sz: 9, color: { rgb: '0000CC' } }
      });

      // Subtotal = PesoTotal × (PUFab + PUMont)
      fmlaCell(8, R, `F${exR}*(G${exR}+H${exR})`, '"S/"#,##0.00', sSubFormula);

      subRows.push(R + 1); // Excel row
      R++;
    });

    // — Fila grupo item (SUM de subitems) —
    const firstSub = subRows[0];
    const lastSub = subRows[subRows.length - 1];
    cell(0, R, `${i + 1}`, sItemGroup);
    merge(1, R, 7, R);
    cell(1, R, it.nombre, sItemGroup);
    for (let c = 2; c <= 7; c++) cell(c, R, '', sItemGroup);
    fmlaCell(8, R, `SUM(I${firstSub}:I${lastSub})`, '"S/"#,##0.00', sItemGroupNum);

    itemTotalRows.push(R + 1);
    R++;
    R++; // espacio entre grupos
  });

  // ══════════════════════════════════════════
  // TOTALES
  // ══════════════════════════════════════════
  // Total KG
  const allSubRows = items.flatMap(it =>
    it.subs.map((_, j) => {
      // reconstruir las filas de subitems
    })
  );

  // Calcular rangos de todas las filas de datos para SUM total KG
  // Más sencillo: recopilar las filas F que tienen subitems
  // Ya que las tenemos en subRows dentro del forEach, hagamos un segundo pass
  // Usamos fórmula que suma toda la columna F entre headerRow+2 y R
  const dataStart = headerRow + 2;
  const dataEnd = R;

  merge(0, R, 4, R);
  cell(0, R, 'TOTAL GENERAL', sTotalLabel);
  for (let c = 1; c <= 4; c++) cell(c, R, '', sTotalLabel);

  // Total KG = SUMA de pesoTotal (col F) de subitems
  // Usamos SUMIF: celdas en col A que contienen "." (son subitems x.y)
  fmlaCell(5, R,
    `SUMPRODUCT((ISNUMBER(FIND(".",A${dataStart}:A${dataEnd})))*F${dataStart}:F${dataEnd})`,
    '#,##0.000 "kg"', sTotal);

  cell(6, R, '', sTotal);
  cell(7, R, '', sTotal);

  // Gran total = suma de los subtotales de grupos de items
  const totalFormula = itemTotalRows.map(r => `I${r}`).join('+');
  fmlaCell(8, R, totalFormula, '"S/"#,##0.00', sTotal);
  R++;

  // ══════════════════════════════════════════
  // Dimensiones hoja
  // ══════════════════════════════════════════
  ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: COLS, r: R } });
  ws['!merges'] = merges;
  ws['!cols'] = [
    { wch: 8 },   // Item
    { wch: 36 },  // Descripción
    { wch: 7 },   // Cant
    { wch: 8 },   // Unidad
    { wch: 13 },  // Peso Unit
    { wch: 13 },  // Peso Total
    { wch: 12 },  // PU Fab
    { wch: 12 },  // PU Mont
    { wch: 14 },  // Subtotal
  ];
  ws['!rows'] = [{ hpt: 28 }]; // fila título más alta

  XLSX.utils.book_append_sheet(wb, ws, 'Cotización');

  // ── Hoja de tarifas (referencia) ──
  const wsTar = {};
  wsTar['A1'] = { v: 'Parámetro', t: 's', s: sHeader };
  wsTar['B1'] = { v: 'Valor', t: 's', s: sHeader };
  wsTar['A2'] = { v: 'Tarifa Fabricación (S//kg)', t: 's' };
  wsTar['B2'] = { v: puFab, t: 'n', z: '"S/"0.00' };
  wsTar['A3'] = { v: 'Tarifa Montaje (S//kg)', t: 's' };
  wsTar['B3'] = { v: puMont, t: 'n', z: '"S/"0.00' };
  wsTar['!ref'] = 'A1:B3';
  wsTar['!cols'] = [{ wch: 28 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, wsTar, 'Tarifas');

  // ── Descargar ──
  const fechaFile = fecha.replace(/\//g, '-');
  XLSX.writeFile(wb, `COMASA_${fechaFile}.xlsx`, { bookType: 'xlsx', cellStyles: true });
  toast('✅ Excel exportado con fórmulas');
}   

// ─── TOAST ───
function toast(msg, err=0) {
  clearTimeout(toastTm);
  const el = document.getElementById('toast');
  el.textContent = err?'✕ '+msg:'✓ '+msg;
  el.className = 'toast show' + (err?' err':'');
  toastTm = setTimeout(() => el.classList.remove('show'), 2600);
}

