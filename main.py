import os
os.environ['USE_PYGEOS'] = '0'
import shapely
import pandas as pd
import geopandas as gpd
import h3pandas
import win32com.client as com

path = os.path.dirname(__file__)

verFileName = 'M2G_CAM_v2.1'
territoryNoToFill = 1
validDateString = '29.03.2023'

def setActiveNetObjects(Visum, validDateString, directedLineUDA):
    
    Visum.Net.Links.SetPassive()
    PT_walkLinks = Visum.Net.Links.FilteredBy('WORDN([TSYSSET],"PTWalk",1)!=[TSYSSET]')
    PT_walkLinks.SetMultipleAttributes(['TSysSet'], [(TSysSet[0]+',Walk',) for TSysSet in PT_walkLinks.GetMultipleAttributes(['TSysSet'])])
    PT_walkLinks = Visum.Net.Links.FilteredBy('WORDN([TSYSSET],"Walk",1)!=[TSYSSET]').SetActive()

    Visum.Net.Nodes.SetPassive()
    Visum.Net.Nodes.FilteredBy('([CountActive:InLinks]>0)&([CountActive:OutLinks]>0)').SetActive()

    Visum.Net.VehicleJourneyItems.SetPassive()
    Visum.Net.VehicleJourneyItems.FilteredBy(f'[IsValid({validDateString})]&([ExtDeparture]>=8.25*3600)&([ExtDeparture]<9.25*3600)').SetActive()

    stopSubset = Visum.Net.Stops.FilteredBy('[Sum:StopAreas\\Sum:StopPoints\\CountActive:ServingVehJourneyItems]>0')

    lineRouteFormulaString = 'WORDN([LineName],\":\",1)+WORDN([LineName],\":\",2)+\":\"+WORDN(NUMTOSTR([EndLineRouteItem\\StopPoint\\StopArea\\Stop\\No]),\".00\",1)'
    Visum.Net.LineRoutes.AddUserDefinedAttribute(directedLineUDA, directedLineUDA, directedLineUDA, 5, formula = lineRouteFormulaString)

    lineAtts = ['Name', 'TSysCode']
    lineTSysDF = pd.DataFrame(Visum.Net.Lines.FilteredBy('[Sum:LineRoutes\\Sum:VehJourneys\\CountActive:VehJourneyItems]>0').GetMultipleAttributes(lineAtts), columns = lineAtts).set_index(lineAtts[0])

    return lineTSysDF, stopSubset

def getHex(Visum, territoryNoToFill, viaCRS = 'WGS84', resolution = 9):
    
    fma = shapely.from_wkt([Visum.Net.Territories.ItemByKey(territoryNoToFill).AttValue('WKTSurface')])
    fmaGDF = gpd.GeoDataFrame(geometry = fma).set_crs(Visum.Net.AttValue('ProjectionDefinition'), allow_override = True).to_crs(viaCRS)
    fmaHex = fmaGDF.h3.polyfill_resample(resolution).to_crs(Visum.Net.AttValue('ProjectionDefinition'))
    fmaHex['centroid'] = fmaHex['geometry'].centroid
    
    return fmaHex

def iterateHex(fmaHex, Visum, mapMatcher, stopSubset, MinIsocTimeAtStopAtt, DirectedLineFreqAtStopAtt, lineTSysDF):

    for n, (i, row) in enumerate(fmaHex.iterrows()):
        
        aTerritory = Visum.Net.AddTerritory(1000 + n, row['centroid'].x, row['centroid'].y)
        aTerritory.SetAttValue('WKTSurface', shapely.to_wkt(row['geometry']))
        
        nearestNode = mapMatcher.GetNearestNode(row['centroid'].x, row['centroid'].y, 960, True)
        
        if nearestNode.Success:
            
            nearestNodeObject = nearestNode.Node
            netElements = Visum.CreateNetElements()
            
            if nearestNodeObject.AttValue('MainNodeNo') == 0:
                netElements.Add(nearestNodeObject)
            
            else:
                netElements.Add(Visum.Net.MainNodes.ItemByKey(nearestNodeObject.AttValue('MainNodeNo')))
            
            fmaHex.loc[i, 'NodeNo'] = nearestNodeObject.AttValue('No')
            fmaHex.loc[i, 'Distance'] = 1.2*nearestNode.Distance
            fmaHex.loc[i, 'Time'] = 0.75*fmaHex.loc[i, 'Distance'] #3600sec/h * dist/m / 4800m/h
            
            Visum.Analysis.Isochrones.ExecutePrT(netElements, 'Walk', 0, 960)
            
            isocString = stopSubset.FilteredBy(f'[{MinIsocTimeAtStopAtt}]<360000000').GetMultipleAttributes([DirectedLineFreqAtStopAtt, MinIsocTimeAtStopAtt])
            isocList = [[valMin.replace('[', '').replace(']', '').split(':') + [valMaj[1]] for valMin in valMaj[0].split('],[')] for valMaj in isocString]
            
            isocFlatDF = pd.DataFrame([item for sublist in isocList for item in sublist], columns = ['Name', 'DirIndicator', 'Freq', 'IsocTime']).set_index('Name')
            isocFlatDF[['Freq', 'IsocTime']] = isocFlatDF[['Freq', 'IsocTime']].astype(int)
            isocFlatDF['IsocTime'] += fmaHex.loc[i, 'Time']
            isocFlatDF = isocFlatDF.join(lineTSysDF, how = 'left')
            
            isocFlatDF_filtered = isocFlatDF[isocFlatDF['IsocTime']<=isocFlatDF['maxWalkTime']].reset_index()
            isocFlatDF_filtered['SWT'] = 1800/isocFlatDF_filtered['Freq'] #0.5 * 60sec/min * 60min/h * Freq/(veh/h)
            isocFlatDF_filtered['AWT'] = isocFlatDF_filtered['SWT'].values + isocFlatDF_filtered['reliabilityFactor']
            isocFlatDF_filtered['TAT'] = isocFlatDF_filtered['IsocTime'] + isocFlatDF_filtered['AWT']
            isocFlatDF_filtered['EDF'] = 1800/isocFlatDF_filtered['TAT']
            
            fmaHex.loc[i, 'AccessIndex'] = 0.5*(isocFlatDF_filtered.groupby(['TSysCode']).max()['EDF'].sum() + isocFlatDF_filtered.groupby(['Name', 'TSysCode']).max()['EDF'].sum())

def main():

    Visum = com.Dispatch('Visum.Visum.230')
    Visum.IO.LoadVersion(os.path.join(path, f'{verFileName}.ver'))

    directedLineUDA = 'DirectedLineName'
    lineTSysDF, stopSubset = setActiveNetObjects(Visum, validDateString, directedLineUDA)

    for i, row in lineTSysDF.iterrows():
        if row['TSysCode'] in ['Bus', 'Ferry', 'Trolleybus']:
            lineTSysDF.loc[i, 'maxWalkTime'] = 8*60
            lineTSysDF.loc[i, 'reliabilityFactor'] = 2*60
        elif row['TSysCode'] in ['Metro', 'Train', 'Tram']:
            lineTSysDF.loc[i, 'maxWalkTime'] = 12*60
            lineTSysDF.loc[i, 'reliabilityFactor'] = 0.75*60
        else:
            print(f"Warning: Behaviour for {row['TSysCode']} not defined.")

    fmaHex = getHex(Visum, territoryNoToFill)

    DirectedLineFreqAtStopAtt = f'Histogram:StopAreas\\Concatenate:StopPoints\\HistogramActive:ServingVehJourneyItems\\VehJourney\\LineRoute\\{directedLineUDA}'

    iterateHex(fmaHex, Visum, Visum.Net.CreateMapMatcher(), stopSubset, 'Min:StopAreas\\Node\\IsocTimePrT', DirectedLineFreqAtStopAtt, lineTSysDF)

    AI_uda = 'AccessIndex'
    Visum.Net.Territories.AddUserDefinedAttribute(AI_uda, AI_uda, AI_uda, 2)
    Visum.Net.Territories.FilteredBy('[No]>=1000').SetMultiAttValues(AI_uda, [(i + 1, val) for i, val in enumerate(fmaHex[AI_uda].values)])

    PTAL_formulaString = f'IF([{AI_uda}]<=0,\"0 (worst)\",IF([{AI_uda}]<=2.5,\"1a\",IF([{AI_uda}]<=5,\"1b\",IF([{AI_uda}]<=10,\"2\",IF([{AI_uda}]<=15,\"3\",IF([{AI_uda}]<=20,\"4\",IF([{AI_uda}]<=25,\"5\",IF([{AI_uda}]<=40,\"6a\",\"6b (best)\"))))))))'
    PTAL_uda = 'PTAL'
    Visum.Net.Territories.AddUserDefinedAttribute(PTAL_uda, PTAL_uda, PTAL_uda, 5, formula = PTAL_formulaString)

    print('done')