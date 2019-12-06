SHEET_FROM_SITE = {
    'BC': 'Temps - Buck Canyon',
    'BKE': 'Temps - Blue Knoll East',
    'DWR': 'Temps - DWR',
    'DWR3': 'Temps - DWR3',
    'HOL': 'Temps - Holiday',
}


def get_temp_sheet_name(site_abbreviation):
    """
    Return the name of the sheet in the excel document containing
    temperature data given the site abbreviation.
    :param site_abbreviation: String abbreviation of the site name.
    :return: String name of the sheet containing temperature data
    for a particular site.
    """
    return SHEET_FROM_SITE[site_abbreviation]
