import numpy as np
from geneticalgorithm import geneticalgorithm as ga

import ExcelInteraction


def f(consumption, roof_catchment, additional_catchment, collection_tank, storage_tank_volume, x, y, tower_height, pump,
      filter_location, one_um_filter, five_um_filter, twenty_um_filter, UV, chemical_type, power_system_choice,
      num_batteries, solar_panel_type, num_solar_panels, inverter, generator):
    return -1 * ExcelInteraction.get_satisfaction(consumption, roof_catchment, additional_catchment, collection_tank,
                                                  storage_tank_volume, x, y, tower_height, pump, filter_location,
                                                  one_um_filter, five_um_filter, twenty_um_filter, UV, chemical_type,
                                                  power_system_choice, num_batteries, solar_panel_type,
                                                  num_solar_panels, inverter, generator)


varbound = np.array([[160, 620],  # consumption
                     [0, 200],  # roof catchment
                     [0, 4]  # collection tank
                     []

                     ])
vartype = np.array([['real'], ['int'], ['int']])

model = ga(function=f, dimension=3, variable_type_mixed=vartype, variable_boundaries=varbound)
model.run()

#  conversion notes:
#