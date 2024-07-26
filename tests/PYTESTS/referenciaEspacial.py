def featureEPSG(cmd):
    # mxd = arcpy.mapping.MapDocument("current")
    # cmds = arcpy.mapping.ListLayers(mxd)
    return arcpy.Describe(cmd).spatialReference.GCSCode
