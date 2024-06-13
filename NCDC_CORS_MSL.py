#
# NCDC_CORS_MSL
#
import subprocess
import pandas as pd 
import geopandas as gpd 
import pymap3d as pm3d
from geodepy import transform
from geodepy.constants import *
from datetime import date, timedelta
from pygeodesy import dms
import simplekml
import calendar
from pathlib import Path
# The reference epoch of ITRF2014 is t0 = 2010.0
# The reference epoch of ITRF2020 is t0 = 2015.0

def FracYear2Date(fr_year):
    ''' convert GPS year e.g. 2021.93 to date(   ) '''
    year,days = divmod( float(fr_year), 1.0 )
    days = days*pd.Timestamp( int(year) , 12, 31).dayofyear
    return date( int(year), 1,1 ) + timedelta( days=days )

class CORS_NCDC:
    def __init__(self):
        Path('./CACHE').mkdir(parents=True, exist_ok=True)
        FILE = r'./Data/Coordinate NCDC ITRF2014@epoch2021.93 for phisan.xlsx'
        df = pd.read_excel( FILE, engine='openpyxl')
        df = df.drop( columns= df.filter(like='Unnamed') )
        df.sort_values( by='STA', ignore_index=True, inplace=True ) 
        df['epoch'] = 'ITRF2014@2021.93'
        def Tr(row):
            ecef = row.X,row.Y,row.Z
            epoch = row.epoch.split('@')[1]
            ecef_tr = transform.conform14(*ecef, FracYear2Date(epoch) , itrf2014_to_itrf2020 )
            geod_tr = pm3d.ecef2geodetic( *ecef_tr[0:3])
            enu = pm3d.ecef2enu( *ecef_tr[0:3], row.Lat, row.Long, row.h, deg=True )
            lat_dms =  dms.toDMS( geod_tr[0] , prec=4 )
            lng_dms =  dms.toDMS( geod_tr[1] , prec=4 )
            return [*geod_tr, lat_dms, lng_dms, *enu]
        df[['lat_', 'lng_','h_', 'lat_dms', 'lng_dms', 'dE','dN','dU' ]] =\
                                df.apply( Tr, axis=1,  result_type='expand' )
        # self.gdf = gpd.GeoDataFrame( df , crs='EPSG:4326',  ... bug!  with pyogrio driver 
        self.gdf = gpd.GeoDataFrame( df , crs='OGC:CRS84', 
                       geometry=gpd.points_from_xy(df.Long, df.Lat) )
        self.InterpolateMSL()

    def InterpolateMSL( self ):
        def getMSL(row):
            CMD = 'GeoidEval --haetomsl -n tgm2017-1  --input-string "{} {} {}"'
            res = subprocess.run( CMD.format(row.Lat,row.Long,row.h), shell=True, capture_output=True )
            msl = float(res.stdout.decode('utf-8').split()[2])
            dlat_,dlng_= row.Lat-row.lat_, row.Long-row.lng_
            return msl,dlat_,dlng_
        self.gdf[['MSL_TGM17','dlat_','dlng_']] =\
                self.gdf.apply( getMSL , axis=1, result_type='expand' )

    def PlotKML( self ):
        kml = simplekml.Kml()
        folder = kml.newfolder(name="NCDC")
        style = simplekml.Style()
        style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/target.png'
        style.iconstyle.color = simplekml.Color.red
        style.labelstyle.scale = 1 
        style.labelstyle.color = simplekml.Color.red
        for idx, row in self.gdf.iterrows():
            #import pdb ; pdb.set_trace()
            desc = f"""
                <html>
                <body>
                <table border="1">
                <tr><th>Parameter</th><th>Value</th></tr>
                <tr><td>Name</td><td>{row['STA']}</td></tr>
                <tr><td>X</td><td>{row['X']:.3f} m</td></tr>
                <tr><td>Y</td><td>{row['Y']:.3f} m</td></tr>
                <tr><td>Z</td><td>{row['Z']:.3f} m</td></tr>
                <tr><td>Lat</td><td>{row['lat_dms']}</td></tr>
                <tr><td>Lng</td><td>{row['lng_dms']}</td></tr>
                <tr><td>HAE</td><td>{row['h']:.3f} m</td></tr>
                <tr><td>MSL_TGM17</td><td>{row['MSL_TGM17']:.3f} m</td></tr>
                <tr><td>Ref.Frame</td><td>{row['epoch']}</td></tr>
                </table>
                </body>
                </html>
                """
            pnt = folder.newpoint(name=f"{idx+1}: {row.STA}", description=desc,
                                  coords=[(row.geometry.x, row.geometry.y)])
            pnt.style = style
        # Save KML to file
        kml.save("./CACHE/CORS_NCDC_MSL.kml")


###############################################################################################
ncdc = CORS_NCDC()
print( ncdc.gdf.dE.describe() )
print( ncdc.gdf.dN.describe() )
print( ncdc.gdf.dU.describe() )
print( ncdc.gdf )

ncdc.PlotKML()  # export to KML , with visualization capability
# simple  KML , need some more preparation
ncdc.gdf.to_file( "./CACHE/CORS_NCDC_MSL_.kml", layer='NCDC_CORS', driver="KML", engine='pyogrio' )

df = ncdc.gdf[[ 'STA', 'X', 'Y', 'Z', 'lat_dms', 'lng_dms',  'h',  'MSL_TGM17','epoch', 'geometry' ]].copy()
df[['h','MSL_TGM17']] = df[['h','MSL_TGM17']].round(3)
df.to_file( "./CACHE/CORS_NCDC_MSL.gpkg", layer='NCDC_CORS', driver="GPKG", engine='pyogrio')

FMT = [ None, None, ',.3f',',.3f', ',.3f', None, None, ',.3f' ]
print(df[df.columns[:-2]].to_markdown(floatfmt=FMT))

df[df.columns[:-2]].to_csv('./CACHE/CORS_NCDC_MSL.csv', index=False ) 

import pdb ; pdb.set_trace()
