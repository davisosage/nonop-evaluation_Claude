# Enverus monthly production (gross 8/8), pulled 2026-11-23 via the Enverus connector; last reported month 2026-08.
# IO 23/26 (Mewbourne, Tonkawa, Roger Mills Co OK). DSU centroid (35.6678, -99.6109) (lateral midpoints measured to it).
# Fit set: 20 mi radius. No Tonkawa well with 12+ months inside 5 or 10 mi; 5 inside 15 and 20 mi (Smith 30X31 #1HB / #2HC, 2018 Diversified, excluded: ~940 lb/ft, peak under 6,000 bbl/month). Strong City 1-30H19 (1.6 mi, 10 months) is a check only.
# api: (well, lateral_ft, first_month, oil_bbl[], gas_mcf[])
WELLS = {
"3512924188": ("SHOW ME THE MONEY #1-16H", 5321, "2025-08",
 [26504,38665,28576,19552,14542,11699,11130,10071,10807,6789,6963,5861],
 [0,82211,77963,60765,45702,43137,39210,53723,45966,35512,34441,36608]),
"3512924133": ("SHOWTIME 1-7H", 5116, "2023-03",
 [5035,28676,24838,20712,16503,12940,10773,8820,6888,6951,5109,5111,4937,4384,3421,3766,3236,2885,2502,2848,3014,2593,2562,1828,1830,2145,1824,1984,1608,1255,884,1631,1271,1417,982,1505,1331,1165,1767,1078,897,1405],
 [0,35877,36521,44017,46105,45508,41205,41366,38257,39680,25881,24970,27720,26450,25388,22962,22506,21167,18844,18107,15403,14433,11912,10323,13556,13205,12836,12204,11517,6011,8169,9832,8692,8425,7128,8276,5290,8328,8363,7544,3914,7512]),
"3512924140": ("SHOWDOWN 1-1H", 5959, "2024-03",
 [5403,22925,17306,13646,10271,7757,6594,5734,4025,4620,4232,2191,3797,2999,2175,2734,2172,1583,1965,2172,1085,2244,1547,2514,2103,1337,2201,1431,1828,186],
 [12151,31098,29440,33132,30468,25339,22104,24083,14515,17558,18008,12922,16347,14906,14319,12828,12501,5668,9547,10519,4092,11832,7521,9164,11185,11068,12842,10839,11617,273]),
"3512924125": ("BIG SHOW USA #1H-1213", 7917, "2022-05",
 [27693,57139,61863,59087,48130,39205,29906,24854,20567,15801,12136,12894,12173,11106,9253,8671,8283,7031,6880,6311,5642,4772,4956,4168,3822,3813,3407,3288,3104,2798,2990,2256,2081,2779,2252,2577,2177,944,1225,1349,1474,1311,1539,1488,1506,1318,1562,1130],
 [36282,69743,75766,87204,96907,112353,109840,112390,103344,89152,33948,38112,41946,43438,42701,43713,49524,50301,59381,65426,63066,50648,54613,53809,46854,42533,38568,35881,37262,34670,33830,31334,25453,29550,28326,28507,25425,108,4230,6795,7415,7231,8383,8342,8937,8790,9049,7632]),
"3512924180": ("BIG SHOW USA 2H-1213", 10098, "2025-09",
 [10612,23327,18773,15954,12265,9824,9101,7489,7135,5742,6028,4644],
 [31068,66008,37192,40191,37326,35219,38683,35566,35124,30939,30362,23003]),
}
LAST_MONTH = "2026-08"

FIT_SET = ['3512924188', '3512924133', '3512924140', '3512924125', '3512924180']   # 20 mi: No Tonkawa well with 12+ months inside 5 or 10 mi; 5 inside 15 and 20 mi (Smith 30X31 #1HB / #2HC, 2018 Diversified, excluded: ~940 lb/ft, peak under 6,000 bbl/month). Strong City 1-30H19 (1.6 mi, 10 months) is a check only
ENV_EUR = {}   # Enverus EUR fields return blank on this subscription; guardrail off

# ProducingDays: Oklahoma rows in Enverus carry calendar days (30/31 every month), so they can't flag a partial
# first month; left empty so build_tc.py uses its month-2 test instead.
DAYS = {}

FIT_SET_10MI = []

# b donor: fewer than 3 fit wells have 24+ months (young Cherokee development). b is fitted on the 24+ month
# Cherokee horizontals in Roger Mills / Custer (Sage Dragon, Mallard 1HX, Keahey 1-27-34 / 1-28-33) and applied to
# this fit set's mean (build_tc.py B_SET). Glass 1-25-24CHX and Troubadour 32-5 2H left out: Aug 2024 gas cliff.
B_SET = []

# Cleaning applied to the raw Enverus rows (make_deal_data.py): interior shut-in months (both streams under 35% of the month before) removed; a single
# stream reported near zero in a month the other stream produced normally (Oklahoma Mar 2026 gas gap) filled with
# the neighbouring months' geometric mean. Excluded from all sets:
#   3512924148 JEANELDA 3/10 BO 1H: negative oil reported in months 1-2
#   3512924142 GLASS 1-25-24CHX: gas drops from 423 to 95 MMcf in one month (Aug 2024) plus shut-ins; reporting artifact
#   3512924144 TROUBADOUR 32-5 2H: same Aug 2024 gas cliff (270 to 130 MMcf)
#   3512924115 SMITH 30X31-16-21 #1HB: 2018 design, peak 6,252 bbl/month
#   3512924116 SMITH 30X31-16-21 #2HC: 2018 design
CLEANING_LOG = [
  '3512924188 2026-04: shut-in month removed (oil 1507, gas 5800)',
  '3512924125 2023-03: shut-in month removed (oil 1426, gas 8314)',
  '3512924125 2025-07: shut-in month removed (oil 0, gas 0)',
  '3512924125 2025-08: shut-in month removed (oil 0, gas 0)',
  '3512924125 2025-09: shut-in month removed (oil 187, gas 0)',
  '3512924125 2022-09: oil 0 filled to 48130',
]
