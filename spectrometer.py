from subprocess import Popen, PIPE
import matplotlib.pyplot as plt
import pandas as pd
import psutil


# Receiving the values of the Y-axis from sampleAPL.exe, analysis and plotting
def real_time_values():
    # The X-axis values are always the same (constants)
    x_axis = get_wavelength_spectrum()
    index = 0
    # Open pipe communication with the SampleAPL.exe
    p = Popen(
        [
            r"C:\Users\amit2\PycharmProjects\untitled2\Minispectrometer_micro "
            r"module_ver1.00.0002\64\DeveloperTools\ForVisualC++\SampleAPL\x64\Debug\SampleAPL.exe"],
        shell=True, stdout=PIPE, stdin=PIPE)
    while True:
        value = str(index) + '\n'
        value = bytes(value, 'UTF-8')  # Needed in Python 3.
        p.stdin.write(value)
        p.stdin.flush()
        hexa_result = p.stdout.readline().strip()  # The values themselves in heaxadecimal - Given in bytes
        print(hexa_result)  # caveman debuging
        hexa_list = hexa_result.split()  # Creating a list from the outputs

        # If the length of the list does not match - this is unnecessary information (usually String).
        if len(hexa_list) < len(x_axis):
            continue
        else:
            index += 1
            # plotting parameters
            y_axis = hex_to_decimal(hexa_list)
            plt.plot(x_axis, y_axis, '-g', linewidth=2, markersize=12)
            plt.ylabel("A/D count")
            plt.xlabel("Wavelength [nm]")
            plt.ylim([0, 4100])
            plt.title('Real-time spectrometer')

            # plotting characteristics
            plt.draw()
            plt.pause(0.01)
            plt.clf()  # clear the board after each


# converting the list from hex to decimal
def hex_to_decimal(hexadecimal_list):
    if hexadecimal_list:
        vector_y = ([int(x, 16) for x in hexadecimal_list])  # converting the list to decimal values
        print(vector_y)
        return vector_y


# Receiving the values of the X-axis
def get_wavelength_spectrum():
    # Insert complete path to the excel file and index of the worksheet (Maybe the excel now it's a little Overkill
    # but meant for Scaleing)
    df = pd.read_excel("./waveLength_consts.xlsx",
                       sheet_name=0)

    x_consts = list(df['waveLength'])
    return x_consts


def main():
    PROCNAME = "SampleAPL.exe"
    real_time_values()
    # Killing the process at the end of the run
    for proc in psutil.process_iter():
        # check the process name matches
        if proc.name() == PROCNAME:
            proc.kill()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)