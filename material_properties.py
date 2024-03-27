# material_properties.py

# Sample material properties data
def get_material_properties(ply_type):
    material_properties = {
        "T300/5208": {"E_x": 181, "E_y": 10.3, "E_s": 7.17, "ν": 0.28},
        "B4/5505": {"E_x": 204, "E_y": 18.5, "E_s": 5.59, "ν": 0.23},
        "AS/H3501": {"E_x": 138, "E_y": 8.96, "E_s": 7.10, "ν": 0.30},
        "Scotchply 1002": {"E_x": 38.6, "E_y": 8.27, "E_s": 4.14, "ν": 0.26},
        "Kevlar49/epoxy": {"E_x": 76, "E_y": 5.50, "E_s": 2.30, "ν": 0.34},
    }

    return material_properties.get(ply_type, {})


def get_strength_properties(ply_type):
    strength_properties = {
        "T300/5208": {"X_t": 1500, "Y_t": 40, "X_c": 1500, "Y_c": 246, "S_c": 68},
        "B4/5505": {"X_t": 1260, "Y_t": 61, "X_c": 2500, "Y_c": 202, "S_c": 67},
        "AS/H3501": {"X_t": 1447, "Y_t": 51.7, "X_c": 1447, "Y_c": 206, "S_c": 93},
        "Scotchply 1002": {"X_t": 1062, "Y_t": 31, "X_c": 610, "Y_c": 118, "S_c": 72},
        "Kevlar49/epoxy": {"X_t": 1400, "Y_t": 12, "X_c": 235, "Y_c": 53, "S_c": 34},
    }

    return strength_properties.get(ply_type, {})
