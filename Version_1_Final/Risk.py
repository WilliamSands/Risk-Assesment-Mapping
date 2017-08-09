import arcpy
from arcpy import env

arcpy.env.overwriteOutput = True
# Set File workspace
file_workspace = 'D:\\PCO.gdb\\Risk_PCO'
env.workspace = file_workspace

# buffer set up
Holdings = 'D:\\PCO.gdb\\Risk_PCO\\Baits_Issued'
Holdings_clip = 'D:\\PCO.gdb\\Risk_PCO\\FARMS_CLIPPED'
distances = [1000, 4000]
unit = "Meters"
# Make a feature layer
arcpy.MakeFeatureLayer_management(Holdings, 'Holdings_Layer')

# searchcursor for evey row in dataset
with arcpy.da.SearchCursor(Holdings, ['Holding_Reference_Number'])as Holdings_Ref_cursor:
    for row in Holdings_Ref_cursor:
        print row[0]
        query = "Holding_Reference_Number = " + str(row[0])
        #print query
        File_output = file_workspace + '\\' 'Buffer_' + str(row[0])
        # print File_output

        # Select Feature using the reference number from holdings layer
        arcpy.SelectLayerByAttribute_management('Holdings_Layer', 'NEW_SELECTION',
                                                "Holding_Reference_Number = " + str(row[0]))



        # Export holding to geodatabase
        Holding_Boundary = file_workspace + '\\' 'Holding_' + str(row[0])
        arcpy.management.CopyFeatures('Holdings_Layer', Holding_Boundary)

        # Mutliple ring Buffer using Selected Features
        arcpy.MultipleRingBuffer_analysis('Holdings_Layer', File_output, distances, unit, "", "ALL")
        arcpy.MakeFeatureLayer_management(File_output, 'Buffer_Layer')
        # arcpy.Buffer_analysis("Holdings_Layer", ofc, var_Buffer, "FULL", "ROUND", "ALL", "")

        #Clip Holdings Buffer



        # Intersect Features
        Intersect_out_features = file_workspace + '\\' 'Intersect_' + str(row[0])
        arcpy.Intersect_analysis([Holdings_clip, File_output], Intersect_out_features, "", "", "INPUT")
        Dissolved_output_intersect = file_workspace + '\\' 'Intersect_Dissolve_' + str(row[0])
        Dissolve_fields_intersect = ['Holding_Name', 'distance']

       # arcpy.Dissolve_management(Intersect_out_features, Dissolved_output_intersect, Dissolve_fields_intersect)
        #print "intersect Complete"

        Clip_output = file_workspace + '\\' 'Clip_' + str(row[0])
        arcpy.Clip_analysis(Holdings_clip, File_output, Clip_output)
        Dissolve_fields = ['Holding_Name']
        Dissolved_output = file_workspace + '\\' 'Dissolve_' + str(row[0])
        # print Dissolved_output
        #arcpy.Dissolve_management(Clip_output, Dissolved_output, Dissolve_fields)



        # Export to Excel

        arcpy.MakeFeatureLayer_management(Dissolved_output_intersect, 'Intersect_Layer')
        arcpy.MakeFeatureLayer_management(Dissolved_output, 'Dissolve_layer')

        Intersect_Selection_Excel = 'distance = 1000'
        #print Intersect_Selection_Excel
        arcpy.SelectLayerByAttribute_management('Intersect_Layer', 'NEW_SELECTION', Intersect_Selection_Excel)
        Excel_Output = 'D:\\PCO\\Map_Output\\'
        Excel_Location = Excel_Output + '\\' + str(row[0]) + '.xls'
        #print Excel_Location
        arcpy.TableToExcel_conversion('Intersect_Layer', Excel_Location)
        # arcpy.SelectLayerByAttribute_management('Intersect_Layer', 'CLEAR_SELECTION')
        print 'Geoprocessing Complete'

        # add Layers to the Map
        mxd = arcpy.mapping.MapDocument(
            'D:\\PCO\\MXD\\Search_Cursor_mxd.mxd')
        df = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
        legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT", "Legend")[0]
        #print legend
        legend.autoAdd = False

        addLayer3 = arcpy.mapping.Layer(Dissolved_output)
        arcpy.ApplySymbologyFromLayer_management(addLayer3,
                                                 'C:\\Database_connections\\assesment_Model_lyr\\Intersect_3.lyr')
        addLayer3.showLabels = True
        arcpy.mapping.AddLayer(df, addLayer3, 'TOP')

        legend.autoAdd = True
        addLayer = arcpy.mapping.Layer(File_output)
        arcpy.ApplySymbologyFromLayer_management(addLayer, 'C:\\Database_connections\\assesment_Model_lyr\\Buffer.lyr')
        arcpy.mapping.AddLayer(df, addLayer, "TOP")

        addLayer2 = arcpy.mapping.Layer(Holding_Boundary)
        arcpy.ApplySymbologyFromLayer_management(addLayer2,
                                                 'C:\\Database_connections\\assesment_Model_lyr\\Holding.lyr')
        arcpy.mapping.AddLayer(df, addLayer2, "TOP")

        legend.autoAdd = False

        #addLayer3 = arcpy.mapping.Layer(Dissolved_output)
        #arcpy.ApplySymbologyFromLayer_management(addLayer3, 'C:\\Database_connections\\assesment_Model_lyr\\Intersect_3.lyr')
        #addLayer3.showLabels = True
        #arcpy.mapping.AddLayer(df, addLayer3, 'TOP')

        arcpy.RefreshActiveView()
        arcpy.RefreshTOC()

        # zoom to layer

        #print df
        lyr = arcpy.mapping.ListLayers(mxd, '', df)[1]
        #print lyr.name
        extent = lyr.getExtent()
        #print extent
        df.extent = extent

        # Adding Title to MXD
        Map_title = "Holding Reference Number : " + ' ' + str(row[0])
        titleItem = arcpy.mapping.ListLayoutElements(mxd, 'TEXT_ELEMENT', '')[0]
        titleItem.text = Map_title


        # Export Map to PNG File
        Png_output = 'D:\\PCO\\Map_Output\\' + str(row[0]) + '.png'
        arcpy.mapping.ExportToPNG(mxd, Png_output)
        #print 'Map Created'
        del mxd

        f = open('B:\\Risk\\Risk.txt', "a")
        f.write(datetime.datetime.now().ctime())  # openlogfile
        f.write('Holding_Reference_Number = '+' ' + str(row[0]) + '\n')
        f.close()
        print "Moving to Next Row"
