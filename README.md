facilityTemperature
===================

Monitor remote temperatures and notify upon deviation

This is a set of asp web pages and vbs scripts designed to be used in conjunction with a mySQL database and a temperature transmitter to allow monitoring temperatures at remote locations.

Requirements (as implemented):
Web server with asp extensions
VK011 temperature sensor interface
Dallas 18s20 temperature sensor (at least 1, interface supports up to 4 and multiple interfaces may be implemented)
LM75 TTL/RS232 converter
Active Experts Active Socket serial driver
mySQL database (with appropriate driver loaded on web server)

The basis of this system is a VK011 temperature sensor interface which uses a Dallas 18s20 temperature sensor and connects to the serial port (rs232) by means of a LM75 TTL/RS232 converter.

In this implementation the serial port is read across the netowrk using a serial socket driver from Active Experts.

The vbs programs are schedules to run on the server and query the sensor interface for the current temperature and then post that data to the mySQL database for later review.

When reading the temperatures the current temp is checked against alarm setpoints (also kept in the mySQL database) and if there is a deviation betond setpoint then a smtp email is sent to the addresses listed in the database for that point.
