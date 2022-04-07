import xlwings as xw


def find_max_satisfaction(sheet):
    max_satisfaction_values = []
    max_satisfaction = 0
    # Consumption
    for consumption in range(300, 620, 200):
        # Roof Catchment
        roof_catchment_options = ["Half Roof", "Full Roof"]
        for roof_catchment in roof_catchment_options:
            # additional catchment
            for additional_catchment in range(0, 200, 100):
                # collection tank
                collection_tank_options = [400, 1500, 2500, 10000]
                for collection_tank in collection_tank_options:
                    # storage tank
                    for storage_tank_volume in range(5, 51, 10):
                        # x
                        for x in range(-20, 100, 10):
                            # y
                            y = 10
                            tower_height = 0
                            # pump
                            pump_options = ["Pump A", "Pump B", "Pump C"]
                            for pump in pump_options:
                                # filter location
                                filter_location_options = ["Up", "Down"]
                                for filter_location in filter_location_options:
                                    # one_um_filter
                                    one_um_filter = 1
                                    for five_um_filter in range(0, 2):
                                        # 20 um filter
                                        for twenty_um_filter in range(0, 2):
                                            # UV
                                            UV_options = ["36W", "50W"]
                                            for UV in UV_options:
                                                # chemical
                                                chemical_options = ["Chlorine", "Ozone"]
                                                for chemical_type in chemical_options:
                                                    # power system
                                                    power_system_options = ["Solar", "Generator"]
                                                    for power_system_choice in power_system_options:
                                                        if power_system_choice == "Solar":
                                                            solar_panel_options = ["HES-260", "SW-80",
                                                                                   "HES-305P"]
                                                            for solar_panel_type in solar_panel_options:
                                                                # solar panel number
                                                                # batteries and num panels
                                                                if power_system_choice == "Solar":
                                                                    num_solar_panels = 0
                                                                    num_batteries = 0
                                                                    if sheet.range('D23') != "":
                                                                        num_solar_panels += 1
                                                                        if sheet.range('D21') != "":
                                                                            num_batteries += 1
                                                                        # inverter
                                                                        inverter = 1
                                                                        generator = 0
                                                                        satisfaction = \
                                                                            get_satisfaction(sheet,
                                                                                             consumption,
                                                                                             roof_catchment,
                                                                                             additional_catchment,
                                                                                             collection_tank,
                                                                                             storage_tank_volume, x,
                                                                                             y, tower_height, pump,
                                                                                             filter_location,
                                                                                             one_um_filter,
                                                                                             five_um_filter,
                                                                                             twenty_um_filter, UV,
                                                                                             chemical_type,
                                                                                             power_system_choice,
                                                                                             num_batteries,
                                                                                             solar_panel_type,
                                                                                             num_solar_panels,
                                                                                             inverter, generator)
                                                                        if isinstance(satisfaction, float):
                                                                            if satisfaction > max_satisfaction:
                                                                                max_satisfaction = satisfaction

                                                                                max_satisfaction_values.append(
                                                                                    consumption)
                                                                                max_satisfaction_values.append(
                                                                                    roof_catchment)
                                                                                max_satisfaction_values.append(
                                                                                    additional_catchment)
                                                                                max_satisfaction_values.append(
                                                                                    collection_tank)
                                                                                max_satisfaction_values.append(
                                                                                    storage_tank_volume)
                                                                                max_satisfaction_values.append(
                                                                                    x)
                                                                                max_satisfaction_values.append(
                                                                                    y)
                                                                                max_satisfaction_values.append(
                                                                                    tower_height)
                                                                                max_satisfaction_values.append(
                                                                                    pump)
                                                                                max_satisfaction_values.append(
                                                                                    filter_location)
                                                                                max_satisfaction_values.append(
                                                                                    one_um_filter)
                                                                                max_satisfaction_values.append(
                                                                                    five_um_filter)
                                                                                max_satisfaction_values.append(
                                                                                    twenty_um_filter)
                                                                                max_satisfaction_values.append(
                                                                                    UV)
                                                                                max_satisfaction_values.append(
                                                                                    chemical_type)
                                                                                max_satisfaction_values.append(
                                                                                    power_system_choice)
                                                                                max_satisfaction_values.append(
                                                                                    num_batteries)
                                                                                max_satisfaction_values.append(
                                                                                    solar_panel_type)
                                                                                max_satisfaction_values.append(
                                                                                    num_solar_panels)
                                                                                max_satisfaction_values.append(
                                                                                    inverter)
                                                                                max_satisfaction_values.append(
                                                                                    generator)
                                                        else:
                                                            num_batteries = 1
                                                            num_solar_panels = 0
                                                            solar_panel_type = "HES-260"
                                                            inverter = 0
                                                            generator = 1

                                                            satisfaction = \
                                                                get_satisfaction(sheet, consumption,
                                                                                 roof_catchment, additional_catchment,
                                                                                 collection_tank, storage_tank_volume,
                                                                                 x, y, tower_height, pump,
                                                                                 filter_location, one_um_filter,
                                                                                 five_um_filter, twenty_um_filter, UV,
                                                                                 chemical_type, power_system_choice,
                                                                                 num_batteries, solar_panel_type,
                                                                                 num_solar_panels, inverter, generator)
                                                            if isinstance(satisfaction, float):
                                                                if satisfaction > max_satisfaction:
                                                                    max_satisfaction = satisfaction

                                                                    max_satisfaction_values.append(consumption)
                                                                    max_satisfaction_values.append(roof_catchment)
                                                                    max_satisfaction_values.append(additional_catchment)
                                                                    max_satisfaction_values.append(collection_tank)
                                                                    max_satisfaction_values.append(storage_tank_volume)
                                                                    max_satisfaction_values.append(x)
                                                                    max_satisfaction_values.append(y)
                                                                    max_satisfaction_values.append(tower_height)
                                                                    max_satisfaction_values.append(pump)
                                                                    max_satisfaction_values.append(filter_location)
                                                                    max_satisfaction_values.append(one_um_filter)
                                                                    max_satisfaction_values.append(five_um_filter)
                                                                    max_satisfaction_values.append(twenty_um_filter)
                                                                    max_satisfaction_values.append(UV)
                                                                    max_satisfaction_values.append(chemical_type)
                                                                    max_satisfaction_values.append(power_system_choice)
                                                                    max_satisfaction_values.append(num_batteries)
                                                                    max_satisfaction_values.append(solar_panel_type)
                                                                    max_satisfaction_values.append(num_solar_panels)
                                                                    max_satisfaction_values.append(inverter)
                                                                    max_satisfaction_values.append(generator)
    return max_satisfaction_values


def get_satisfaction(sheet, consumption, roof_catchment, additional_catchment, collection_tank, storage_tank_volume, x, y,
                     tower_height, pump, filter_location, one_um_filter, five_um_filter, twenty_um_filter, UV,
                     chemical_type, power_system_choice, num_batteries, solar_panel_type, num_solar_panels, inverter,
                     generator):
    sheet.range('B2').value = consumption
    sheet.range('C5').value = roof_catchment
    sheet.range('C6').value = additional_catchment
    sheet.range('C7').value = collection_tank
    sheet.range('C8').value = storage_tank_volume
    sheet.range('C9').value = x
    sheet.range('C10').value = y
    sheet.range('C12').value = tower_height
    sheet.range('C13').value = pump
    sheet.range('C14').value = filter_location
    sheet.range('C15').value = one_um_filter
    sheet.range('C16').value = five_um_filter
    sheet.range('C17').value = twenty_um_filter
    sheet.range('C18').value = UV
    sheet.range('C19').value = chemical_type
    sheet.range('C20').value = power_system_choice
    sheet.range('C21').value = num_batteries
    sheet.range('C22').value = solar_panel_type
    sheet.range('C23').value = num_solar_panels
    sheet.range('C24').value = inverter
    sheet.range('C25').value = generator

    if (sheet.range('K5') != "Requirement not met" and sheet.range('K6') != "Requirement not met" and
            sheet.range('K7') != "Requirement not met" and sheet.range('K8') != "Requirement not met" and
            sheet.range('K9') != "Requirement not met" and sheet.range('K10') != "Requirement not met" and
            sheet.range('K11') != "Requirement not met"):

        return sheet.range('F14').value
    else:
        return 0


def main():
    wb = xw.Book('RWHSimCopy.xlsx')  # connect to a file that is open or in the current working directory

    sheet = wb.sheets['Parameters and Satisfaction']
    find_max_satisfaction(sheet)


main()
