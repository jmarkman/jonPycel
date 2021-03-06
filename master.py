import sovinput
import sovmanip as modify
import os
import uuid


comparisonDic = {
    'Yr Bldg updated (Mand if >25 yrs)': 'Wiring Year',
    'name/address': 'Street 1',
    'loc': 'Loc #',
    'clt #': 'Loc #',
    'Loc  #': 'Loc #',
    'Other Value $ (outdoor prop & Eqpt must be sch\'d)': 'Business Income',
    'year built': 'Year Built',
    'yearbuilt(yyyy)': 'Year Built',
    'yr built': 'Year Built',
    'sq footage': 'Sq Ftg',
    'totalsqft': 'Sq Ftg',
    'location(s)': 'Street 1',
    'building address': 'Street 1',
    'loc. #': 'Loc #',
    'bldg.': 'Bldg #',
    'bldg limit': 'Building Value',
    'business income/ lor': 'Business Income',
    'bpp w/stock': 'Business Personal Property',
    'ocupancy': 'Building Description',
    '*real property value ($)': 'Building Value',
    'real property value': 'Building Value/',
    'no. of stories': '# Stories',
    'storiesaboveground': '# Stories',
    'elev': '# Stories',
    'building square footage': 'Sq Ftg',
    'grosssqfootage': 'Sq Ftg',
    'bldg': 'Bldg #',
    'code': 'ClassCodeDesc',
    'const code (iso)*': 'Construction Type',
    'no. of units': '# Units',
    'roof shape': 'Roof Geometry',
    '* bi/rental income': 'Business Income',
    '* contents value': 'Business Personal Property',
    'bi w/ee': 'Business Income',
    'building num .': 'Bldg #',
    'sprinkler extent': 'Sprinkler Extent',
    '# of units': '# Units',
    'unit #': '# Units',
    '*total tiv': 'TIV',
    'fire alarm type': 'Fire Alarm Type',
    'fire alarm': 'Fire Alarm Type',
    '*bldg no.': 'Bldg #',
    'street add.': 'Street 1',
    'loss of bus iness income/rents @ 100% annual': 'Business Income',
    '*yr. roof covering last repl': 'Roofing Year',
    'building description': 'Building Description',
    'const': 'Construction Type',
    'constr': 'Construction Type',
    'const floors': 'Construction Type',
    'street address': 'Street 1',
    'street name & number': 'Street 1',
    'building square ft.': 'Sq Ftg',
    'cnst type': 'Construction Type',
    'plumbing updates': 'Plumbing Year',
    'location': 'Loc #',
    'bi/ee': 'Business Income',
    '*# of units if apartments or condos': '# Units',
    'sq. ft.': 'Sq Ftg',
    'protect class': 'Prot Class',
    'address:': 'Street 1',
    'year roof updated': 'Roofing Year',
    'street': 'Street 1',
    '**exterior cladding': 'Cladding Type',
    '* bldg no.': 'Bldg #',
    'address1': 'Street 1',
    'building number': 'Bldg #',
    '# of sto.': '# Stories',
    'contents value': 'Business Personal Property',
    'soft costs': 'Business Personal Property',
    'insured contact address': 'Street 1',
    'constr *': 'Construction Type',
    'sprinklered y or n': 'Sprinkler Alarm Type',
    'county': 'County',
    'street 1': 'Street 1',
    'burgler alarm': 'Burgler Alarm Type',
    'electrical update yr': 'Wiring Year',
    'yearelectricalupdated(yyyy)': 'Wiring Year',
    'yr blt': 'Year Built',
    'bpp': 'Business Personal Property',
    'yr of updates to wiring': 'Wiring Year',
    'loc num': 'Loc #',
    'bldg. num.': 'Bldg #',
    'business income & extra expen se': 'Business Income',
    'exterior cladding': 'Cladding Type',
    'sqfootage': 'Sq Ftg',
    'rents': 'Business Income',
    '# st': '# Stories',
    'bld#': 'Bldg #',
    'insured contace zipcode': 'Zip',
    'const.': 'Con struction Type',
    'wiring updates': 'Wiring Year',
    'sprinkler alarm type': 'Sprinkler Alarm Type',
    'y r. built': 'Year Built',
    'loc id': 'Loc #',
    'roof': 'Roofing Year',
    'year': 'Year Built',
    'bpp (cont ents)     *this is not included in the buildings coverage.': 'Business Personal Property',
    '*zip': 'Zip',
    'zip': 'Zip',
    'misc real property': 'Misc Real Property',
    'bi limit': 'Business Income',
    'total building squ are footage': 'Sq Ftg',
    'roofing year': 'Roofing Year',
    'total': 'TIV',
    'prot. class': 'Prot Class',
    'predominant exterior wall / cladding (use weakest cladding comprising at least 25% of wall area)': 'Cladding Type',
    'fire alarms (operational)': 'Fire Alarm Type',
    'totals': 'TIV',
    'bldg.': 'Bldg #',
    'year electric updated': 'Wiring Year',
    'state': 'State',
    'protections': 'Prot Class',
    'type of ro of covering': 'Roof Covering',
    'building no.': 'Bldg #',
    'burgler alarm type': 'Burgler Alarm Type',
    'stories': '# Stories',
    'frame-foundation connection': 'Frame-Foundation Connection',
    'insureds com plete street address': 'Street 1',
    'percent sprinklered': 'Sprinkler Extent',
    'em cladding type': 'Cladding Type',
    'bldg sq ft': 'Sq Ftg',
    'pers prop': 'Business Personal Property',
    '*street address': 'Street 1',
    'protection class': 'Prot Class',
    'bi': 'Business Income',
    '# basements': '# Basements',
    '# of bldgs': '# Units',
    'total insurable values': 'TIV',
    'address': 'Street 1',
    'building frame to foundation connection': 'Frame-Foundation Connection',
    'pl umbing update yr': 'Plumbing Year',
    'miscelaneous real property': 'Misc Real Property',
    'square feet ': 'Sq Ftg',
    '*# of stories': '# Stories',
    '*roof anchorage (if iso 1 or 2 or any other with wood fr amed roof)': 'Roof Anchor',
    'zip code': 'Zip',
    '*square footage': 'Sq Ftg',
    'personal property value ($)': 'Business Personal Property',
    'roof covering': 'Roof Covering',
    'loc #': 'Loc #',
    'aplocnumbe r': 'Loc #',
    'loc no.': 'Loc #',
    'total sf': 'Sq Ftg',
    '*# of bldgs': '# Units',
    'property type': 'Building Description',
    'occ': 'Building Description',
    'physical address': 'Street 1',
    'professional s qft': 'Sq Ftg',
    'construction description': 'Construction Type',
    '# stories': '# Stories',
    'occupanc y / building type': 'Building Description',
    '*state': 'State',
    'bld #': 'Bldg #',
    '*real property va lue ($)': 'Building Value',
    'plumbing year': 'Plumbing Year',
    'yearplumbingupdated(yyyy)': 'Plumbing Years',
    'rating basis': 'Building Description',
    'business income/ rents/ inc extra expense': 'Business Income',
    'construction description (provide further details on construction features)': 'Construction Type',
    'roof geometry': 'Roof Geometry',
    'insured contact state': 'State',
    'sq ftg': 'Sq Ftg',
    'location number': 'Loc #',
    "location #": 'Loc #',
    'city': 'City',
    'zip': 'Zip',
    'area': 'Sq Ftg',
    'business income': 'Business Income',
    'st.': 'Street 1',
    'loc number': 'Loc #',
    'construction description (ie frame, jm, nc, mnc, fire resistive, modified fire resistive, etc)': 'Construction Type',
    'occupancy': 'Building Description',
    'bldggsqft': 'Sq Ftg',
    'pc': 'Prot Class',
    'roof anchor': 'Roof Anchor',
    'postal code': 'Zip',
    '*state abbrev.': 'State',
    'square ft.': 'Sq Ftg',
    'roof wall attachment': 'Roof Sheathing Attachment',
    'bldg value': 'Building Value',
    '% sprkld': 'Sprinkler Extent',
    'yr of updates to plumbing': 'Plumbing Year',
    'construction': 'Construction Type',
    'classcodedesc': 'ClassCodeDesc',
    '*city': 'City',
    'iso prot class': 'Prot Class',
    'isoconstcode': 'Prot Class',
    'tota l values': 'TIV',
    'location address': 'Street 1',
    'address including street #': 'Street 1',
    'yr of u pdates to roofing': 'Roofing Year',
    'sqft': 'Sq Ftg',
    'cladding type': 'Cladding Type',
    '* building value': 'Building Value',
    'building value': 'Building Value',
    'bi/rental income ($)': 'Misc Real Property',
    '**shape of roof': 'Roof Geometry ',
    'loss of business income': 'Business Income',
    'hard cost': 'Building Value',
    '*basement': '# Basements',
    'extshell': 'Construction Type',
    'wiring year': 'Wiring Year',
    '*occupancy description': 'Building Description',
    'roof update yr': 'Roofing Year',
    'yearroofupdated(yyyy)': 'Roofing Year',
    'annual rents': 'Business Income',
    'rents value': 'Business Income',
    'rents 100% (12 months)': 'Business Income',
    '**type of roof covering': 'Roof Geometry',
    'prot class': 'Prot Class',
    "# bldg's": '# Units',
    'sprlk': 'Sprinkler Alarm Type',
    'wh at type of construction is the building?   (see yellow second tab for descriptions)': 'Construction Type',
    'constructionclassdescription(isoseeattached)': 'Construction Type',
    'state/province': 'State',
    'building replacement cost': 'Building Value',
    '*roof anchorage': 'Roof Anchor',
    'construction type': 'Construction Type',
    'constru-ctiondescription': 'Construction Type',
    'year plumbing updated': 'Plumbing Year',
    'building limit': 'Building Value',
    'sq. foot': 'Sq Ftg',
    'location city': 'City',
    'contents': 'Business Personal Property',
    'all other personal property': 'Business Personal Property',
    'personal prope rty': 'Business Personal Property',
    '*county': 'County',
    'number of stories': '# Stories',
    'sprinkle r': 'Sprinkler Alarm Type',
    '*occupancy type (ie apartments, offices, warehouses)': 'Building Descri ption',
    'building @ 100%': 'Building Value',
    'tiv': 'TIV',
    'wiring': 'Wiring Year',
    'zipcode': 'Zip',
    'personal property limit': 'Business Personal Property',
    '(tiv) total insurable value': 'TIV',
    'un its': '# Units', '# of floors': '# Stories',
    'bus income': 'Business Income',
    'type of construction': 'Construction Type',
    'insured contact city': 'City',
    '# units': '# Units',
    'roof updates': 'Roofin g Year',
    'square footage': 'Sq Ftg',
    'business personal property': 'Business Personal Property',
    'ye ar built': 'Year Built',
    'total above ground sqft': 'Sq Ftg',
    'shape of roof': 'Roof Geometry',
    'yea r roof covering last fully replaced': 'Roofing Year',
    'basement': '# Basements',
    '*property type': 'Building Description',
    'ppc code': 'Prot Class',
    'building': 'Building Value',
    'const type': 'Construction Type',
    'state code': 'State',
    '*state code': 'State',
    'sprinklered (y/n)': 'Sprinkler Alarm Type',
    'sprinkler system ': 'Sprinkler Alarm Type',
    'bldgname': 'Street 1',
    'description': 'Building Description',
    'st': 'State',
    '*orig year built': 'Year Built',
    'year built': 'Year Built',
    'bi/ee/rents': 'Business Income',
    'plumbing': 'Plumbing Year',
    'sq ft': 'Sq Ftg',
    'sf': 'Sq Ftg',
    'yearhvacupdated(yyyy)': 'Heating Year',
    'heatingyear': 'Heating Year',
    'Location Street Address: ': 'Street 1'
}
userhome = os.path.expanduser('~/Desktop/')

clear = lambda: os.system('cls')
clear()

input_sov = sovinput.ask()
sov_sheet = sovinput.findSheetName(input_sov)
sov_data = sovinput.loopAllRows(sov_sheet)
header_row = sovinput.identifyHeaderRow(sov_data, comparisonDic)

header_row = modify.head_matcher(
modify.comp_converter(comparisonDic), header_row, input_sov)
sub_header_data = modify.sliceSubHeaderData(header_row, sov_sheet)
head_sub_combine = modify.combine(header_row, sub_header_data)
modify.setnwrite(head_sub_combine, input_sov)


