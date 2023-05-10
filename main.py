from invester.models.robot import FredData

if __name__ == '__main__':
    fred_data = FredData('[my_api_key]')
    series_names = ['FEDFUNDS', 'DGS10', 'T10YFF', 'BAA10Y', 'TWEXBGSMTH']
    fred_data.get_latest_data(series_names)
    fred_data.get_last_year_data(series_names)
    cell_locations = {
        'FEDFUNDS': {'latest': 'B6', 'last_year': 'D6'},
        'DGS10': {'latest': 'B7', 'last_year': 'D7'},
        'T10YFF': {'latest': 'B8', 'last_year': 'D8'},
        'BAA10Y': {'latest': 'B9', 'last_year': 'D9'},
        'TWEXBGSMTH': {'latest': 'B11', 'last_year': 'D11'}
    }
    fred_data.write_to_excel(r'C:\Users\user\Documents\Documents\investment_environment.xlsx', cell_locations)