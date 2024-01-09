import ifcopenshell
import pandas as pd
import os
from tkinter import Tk, filedialog
from datetime import datetime

def extract_space_details(ifc_file):
    ifc_model = ifcopenshell.open(ifc_file)
    spaces = ifc_model.by_type('IfcSpace')
    data = []

    # Extrahiere den Dateinamen aus dem Pfad
    file_name = os.path.basename(ifc_file)

    # Extrahiere den Timestamp aus der IFC-Datei
    timestamp_ifc = None
    owner_history = ifc_model.by_type('IfcOwnerHistory')
    if owner_history:
        timestamp_ifc = owner_history[0].CreationDate
        timestamp_ifc = datetime.utcfromtimestamp(timestamp_ifc).strftime("%Y-%m-%d %H:%M")

    # Exportdatum erstellen
    export_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

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

        data.append([timestamp_ifc, export_timestamp, file_name, name, long_name, net_floor_area, net_perimeter])

    return pd.DataFrame(data, columns=['Timestamp_IFC', 'ExportTimestamp', 'FileName', 'Name', 'LongName', 'NetFloorArea', 'NetPerimeter'])

def main():
    root = Tk()
    root.withdraw()  # Versteckt das Hauptfenster von Tkinter

    file_path = filedialog.askopenfilename(title="Wähle eine IFC-Datei aus")
    if file_path:
        df = extract_space_details(file_path)
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        output_file = os.path.join(desktop, 'IFC_Space_Details.xlsx')
        df.to_excel(output_file, index=False, columns=['Timestamp_IFC', 'ExportTimestamp', 'FileName', 'Name', 'LongName', 'NetFloorArea', 'NetPerimeter'])
        print(f"Die Details wurden in {output_file} gespeichert.")

if __name__ == '__main__':
    main()
