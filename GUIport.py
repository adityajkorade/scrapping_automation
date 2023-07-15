import serial
import matplotlib.pyplot as plt

# Define serial port parameters
port = 'COM 5'  # Update with your Arduino port
baud_rate = 9600

# Define data storage variables
data = []  # List to store received data

# Create a figure and axis for live chart
fig, ax = plt.subplots()
plt.ion()  # Enable interactive mode

# Open serial port
ser = serial.Serial(port, baud_rate)

# Read and plot data in a loop
while True:
    # Read data from serial port
    line = ser.readline().decode().strip()  # Assuming text-based data

    # Split the received line into values (modify this based on your data format)
    values = line.split(',')

    # Convert values to appropriate data types if needed
    # Example: value1 = int(values[0])
    #         value2 = float(values[1])

    # Store data in a list
    data.append(values)

    # Update live chart
    # Modify this section to plot the specific data you want
    # Example: ax.plot(range(len(data)), [row[0] for row in data], 'b-')
    ax.plot(range(len(data)), [row[1] for row in data], 'r-')
    plt.draw()
    plt.pause(0.01)  # Pause to allow the plot to update

    # Save data to a file
    with open('data.txt', 'w') as file:
        for row in data:
            file.write(','.join(row) + '\n')

    # Break the loop if needed
    # Example: if condition_met: break

# Close the serial port
ser.close()
