import numpy as np

import datetime
from datetime import date, timedelta

import sqlalchemy
from sqlalchemy.ext.automap import automap_base
from sqlalchemy.orm import Session
from sqlalchemy import create_engine, func

from flask import Flask, jsonify


#################################################
# Database Setup
#################################################
engine = create_engine("sqlite:///hawaii.sqlite")

# reflect an existing database into a new model
Base = automap_base()
# reflect the tables
Base.prepare(engine, reflect=True)

# Save reference to the table
Measurement = Base.classes.measurement
Station = Base.classes.station

# Create our session (link) from Python to the DB
session = Session(engine)

#################################################
# Flask Setup
#################################################
app = Flask(__name__)

last_date = datetime.datetime.strptime("2017-08-23", "%Y-%m-%d")
date_list = [last_date - datetime.timedelta(days=x) for x in range(0, 365)]

date_data = []
for date in date_list:
    new_date = date.strftime("%Y-%m-%d")
    date_data.append(new_date)

#################################################
# Flask Routes
#################################################

@app.route("/")
def welcome():
    """List all available api routes."""
    return (
        f"Available Routes:<br/>"
        f"/api/v1.0/precipitation<br/>"
        f"/api/v1.0/stations<br/>"
        f"/api/v1.0/tobs<br/>"
        f"/api/v1.0/<start><br/>"
        f"/api/v1.0/<start>/<end><br/>"
        f"Please enter dates in the 'YYYY-MM-DD' format."
    )

@app.route("/api/v1.0/precipitation")
def precipitation():

    results = session.query(Measurement).filter(Measurement.date.in_(date_data))
    
    precipitation_data = []
    for day in results:
        precipitation_dict = {}
        precipitation_dict[day.date] = day.prcp
        precipitation_data.append(precipitation_dict)

    return jsonify(precipitation_data)

@app.route("/api/v1.0/stations")
def stations():

    results = session.query(Station)

    station_data = []
    for station in results:
        station_dict = {}
        station_dict["Station"] = station.station
        station_dict["Name"] = station.name
        station_data.append(station_dict)

    return jsonify(station_data)

@app.route("/api/v1.0/tobs")
def tobs():

    results = session.query(Measurement).filter(Measurement.date.in_(date_data))

    temperature_data = []
    for day in results:
        temperature_dict = {}
        temperature_dict[day.date] = day.tobs
        temperature_data.append(temperature_dict)

    return jsonify(temperature_data)

@app.route("/api/v1.0/<start>")
def temperature_start(start):
    
    start_date = datetime.datetime.strptime(start, "%Y-%m-%d")
    end_date = datetime.datetime.strptime("2017-08-23", "%Y-%m-%d")

    delta = end_date - start_date
    date_range = []
    for i in range(delta.days + 1):
        date_range.append(start_date + timedelta(days=i))
    
    date_range_list = []
    for date in date_range:
        new_date = date.strftime("%Y-%m-%d")
        date_range_list.append(new_date)
    
    min_temperature = session.query(func.min(Measurement.tobs))\
                            .filter(Measurement.date.in_(date_range_list))[0][0]
    max_temperature = session.query(func.max(Measurement.tobs))\
                            .filter(Measurement.date.in_(date_range_list))[0][0]
    avg_temperature = session.query(func.avg(Measurement.tobs))\
                            .filter(Measurement.date.in_(date_range_list))[0][0]

    temperature_dict = {}
    temperature_dict["Lowest Temperature"] = min_temperature
    temperature_dict["Highest Temperature"] = max_temperature
    temperature_dict["Average Temperature"] = avg_temperature

    return jsonify(temperature_dict)

@app.route("/api/v1.0/<start>/<end>")
def temperature(start, end):

    start_date = datetime.datetime.strptime(start, "%Y-%m-%d")
    end_date = datetime.datetime.strptime(end, "%Y-%m-%d")

    delta = end_date - start_date
    date_range = []
    for i in range(delta.days + 1):
        date_range.append(start_date + timedelta(days=i))
    
    date_range_list = []
    for date in date_range:
        new_date = date.strftime("%Y-%m-%d")
        date_range_list.append(new_date)
    
    min_temperature = session.query(func.min(Measurement.tobs))\
                            .filter(Measurement.date.in_(date_range_list))[0][0]
    max_temperature = session.query(func.max(Measurement.tobs))\
                            .filter(Measurement.date.in_(date_range_list))[0][0]
    avg_temperature = session.query(func.avg(Measurement.tobs))\
                            .filter(Measurement.date.in_(date_range_list))[0][0]

    temperature_dict = {}
    temperature_dict["Lowest Temperature"] = min_temperature
    temperature_dict["Highest Temperature"] = max_temperature
    temperature_dict["Average Temperature"] = avg_temperature

    return jsonify(temperature_dict)

if __name__ == '__main__':
    app.run(debug=True)