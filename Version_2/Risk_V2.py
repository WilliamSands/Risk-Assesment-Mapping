import arcpy
from arcpy import env
import time

# Settings
baseDirectory = 'C:\\Users\\hawkinle\\Desktop\\files'
file_workspace = baseDirectory + '\\Leith.gdb'
Holdings = file_workspace + '\\Holdings_Join'
Holdings_clip = file_workspace + '\\Holdings'
distances = [1000, 4000]
unit = "Meters"

# # ArcPy settings
arcpy.env.overwriteOutput = True
env.workspace = file_workspace

arcpy.MakeFeatureLayer_management(Holdings, 'Holdings_Layer')

mxd = arcpy.mapping.MapDocument(baseDirectory + '\\Search_Cursor_mxd.mxd')
df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT", "Legend")[0]
legend.autoAdd = True


def addLayerToMxd(layer_location, lyr_file_location, show_labels):
    layer = arcpy.mapping.Layer(layer_location)
    arcpy.ApplySymbologyFromLayer_management(layer, baseDirectory + lyr_file_location)
    layer.showLabels = show_labels
    arcpy.mapping.AddLayer(df, layer, 'TOP')


with arcpy.da.SearchCursor(Holdings, ['Holding_Reference_Number'])as Holdings_Ref_cursor:
    for row in Holdings_Ref_cursor:
        start = time.time()
        refNumber = str(row[0])
        print 'Holding:' + refNumber

        # Select the holding (eg a definition query)
        arcpy.Select_analysis('Holdings_Layer', 'in_memory/holding', "Holding_Reference_Number = " + refNumber)

        arcpy.Buffer_analysis('in_memory/holding', 'in_memory/buffer1km', "1000 Meters", "FULL", "ROUND", "ALL")
        arcpy.AddField_management("in_memory/buffer1km", "distance", "SHORT")
        arcpy.CalculateField_management("in_memory/buffer1km", "distance", 1)

        arcpy.Buffer_analysis('in_memory/holding', 'in_memory/buffer4km', "4000 Meters", "FULL", "ROUND", "ALL")
        arcpy.AddField_management("in_memory/buffer4km", "distance", "SHORT")
        arcpy.CalculateField_management("in_memory/buffer4km", "distance", 4)

        arcpy.Merge_management(["in_memory/buffer1km", "in_memory/buffer4km"], "in_memory/buffer")


        arcpy.Intersect_analysis([Holdings_clip, 'in_memory/buffer'], 'in_memory/intersect', "", "", "INPUT")

        arcpy.Dissolve_management('in_memory/intersect', 'in_memory/intersectDissolved', ['Holding_Name', 'distance'])

        arcpy.Clip_analysis(Holdings_clip, 'in_memory/buffer', 'in_memory/clip')

        arcpy.Dissolve_management('in_memory/clip', 'in_memory/dissolvedOutput', ['Holding_Name'])
        # Export to Excel
        arcpy.MakeFeatureLayer_management('in_memory/intersectDissolved', 'Intersect_Layer')

        arcpy.SelectLayerByAttribute_management('Intersect_Layer', 'NEW_SELECTION', 'distance=1000')

        arcpy.TableToExcel_conversion('Intersect_Layer', baseDirectory + '\\blah' + refNumber + '.xls')

        # Add the layers to the mxd
        addLayerToMxd('in_memory/dissolvedOutput', '\\Intersect.lyr', True)
        addLayerToMxd('in_memory/buffer', '\\Buffer.lyr', False)
        addLayerToMxd('in_memory/holding', '\\Holding.lyr', False)

        arcpy.RefreshActiveView()
        arcpy.RefreshTOC()

        # Zoom to the extent of the holding
        lyr = arcpy.mapping.ListLayers(mxd, '', df)[1]
        df.extent = lyr.getExtent()

        # Adding Title to MXD
        titleItem = arcpy.mapping.ListLayoutElements(mxd, 'TEXT_ELEMENT', '')[0]
        titleItem.text = "Holding Reference Number : " + ' ' + refNumber

        # Export Map to PNG File
        arcpy.mapping.ExportToPNG(mxd, baseDirectory + '\\' + refNumber + '.png')
print('Time: ' + str(time.time() - start))
