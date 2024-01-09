import ifcopenshell
import pandas as pd
import os
from tkinter import Tk, filedialog

def extract_space_details(ifc_file):
    ifc_model = ifcopenshell.open(ifc_file)
    spaces = ifc_model.by_type('IfcSpace')
    data = []

    for space in spaces:
        net_floor_area = None
        net_perimeter = None
        name = space.Name or ''
        long_name = space.LongName or ''

        for relDefines in space.IsDefinedBy:
            if relDefines.is_a('IfcRelDefinesByProperties'):
                property_set = relDefines.RelatingPropertyDefinition
                if property_set.is_a('IfcElementQuantity'):
                    if property_set.Name == 'BaseQuantities':
                        for quantity in property_set.Quantities:
                            if quantity.Name == 'NetFloorArea':
                                net_floor_area = quantity.AreaValue
                            elif quantity.Name == 'NetPerimeter':
                                net_perimeter = quantity.LengthValue

        data.append([space.GlobalId, name, long_name, net_floor_area, net_perimeter])

    return pd.DataFrame(data, columns=['GlobalId', 'Name', 'LongName', 'NetFloorArea', 'NetPerimeter'])

def main():
    root = Tk()
    root.withdraw()  # Versteckt das Hauptfenster von Tkinter

    file_path = filedialog.askopenfilename(title="Wähle eine IFC-Datei aus")
    if file_path:
        df = extract_space_details(file_path)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        output_file = os.path.join(desktop, 'IFC_Space_Details.xlsx')
        df.to_excel(output_file, index=False)
        print(f"Die Details wurden in {output_file} gespeichert.")

if __name__ == '__main__':
    main()
