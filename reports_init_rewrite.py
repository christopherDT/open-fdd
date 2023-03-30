import pandas as pd
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import os
from docx import Document
from docx.shared import Inches
import math
import os
import time
from io import BytesIO


FAULT_COLS_DICT = {
    1: ['duct_static', 'supply_vfd_speed', 'duct_static_setpoint'],
    2: ['mat', 'rat', 'oat'],
    # 2: ['mixing_air_temperature', 'return_air_temperature', 'outside_air_temperature', 'supply_vfd_speed'],
    9: ['satsp', 'oat', 'economizer_sig', 'cooling_sig', 'supply_vfd_speed']
}

class ReportCalculator:
    def __init__(self, fault_num, df):

        self.fault_num = fault_num
        self.fault_cols = FAULT_COLS_DICT[fault_num]
        self.df = df


    def summarize_fault_times(self, output_col: str = None) -> str:
        df = self.df

        # if output_col is None:
        output_col = f"fc{self.fault_num}_flag"

        delta = df.index.to_series().diff()

        total_days = round(delta.sum() / pd.Timedelta(days=1), 2)

        total_hours = delta.sum() / pd.Timedelta(hours=1)

        hours_fault_mode = (delta * df[output_col]).sum() / pd.Timedelta(hours=1)

        percent_in_fault_mode = round(df[output_col].mean() * 100, 2)

        #  flag_true_mat = round(
        #     df[self.mat_col].where(df[output_col] == 1).mean(), 2
        # )
        # then later:             paragraph.add_run(
                # f'When fault condition 5 is True the average mix air temp is {flag_true_mat}°F and the outside air temp is {flag_true_sat}°F. This could possibly help with pin pointing AHU operating conditions for when this fault is True.')

        avg_fault_cond_vals = {}
        for col in self.fault_cols:
            avg_fault_cond_vals[col] = round(df[col].where(df[output_col] == 1).mean(), 2)

        motor_on = df['supply_vfd_speed'].gt(1.).astype(int)
        hours_motor_runtime = round(
            (delta * motor_on).sum() / pd.Timedelta(hours=1), 2)

        # for summary stats on I/O data to make useful
        df_motor_on_filtered = df[df['supply_vfd_speed'] > 1.0]
        
        return (
            total_days,
            total_hours,
            hours_fault_mode,
            percent_in_fault_mode,
            avg_fault_cond_vals,
            hours_motor_runtime,
            df_motor_on_filtered
        )



# fault calculator just generates a flag for when the fault is true or not (simple single boolean column)
    # then report generator takes that and can calculate summary stats, etc.

# each fault has:
#   - a plain text definition of the fault condition
#   - a calculation for its fault condition
#       - the required sensors/variables for that calculation
#       - a plain text name for each sensor
#   - different dataset plots

#   - a plain text description of the avg values for each fault condition sensor (e.g. "when fault condition 2 was Treu, avg mat was X degrees F")

FAULT_DEFS_DF = pd.DataFrame(
    [
        [1, ['duct_static', 'supply_vfd_speed', 'duct_static_setpoint'], """Fault condition one of ASHRAE Guideline 36 is related to flagging poor performance of a AHU variable supply fan attempting to control to a duct pressure setpoint. Fault condition equation as defined by ASHRAE:"""],
        [2, ['mat', 'rat', 'oat'], """Fault condition two and three of ASHRAE Guideline 36 is related to flagging mixing air temperatures of the AHU that are out of acceptable ranges. Fault condition 2 flags mixing air temperatures that are too low and fault condition 3 flags mixing temperatures that are too high when in comparision to return and outside air data. The mixing air temperatures in theory should always be in between the return and outside air temperatures ranges. Fault condition two equation as defined by ASHRAE:"""],
        [9, ['satsp', 'oat', 'economizer_sig', 'cooling_sig', 'supply_vfd_speed'], """Fault condition nine of ASHRAE Guideline 36 is an AHU economizer free cooling mode only with an attempt at flagging conditions where the outside air temperature is too warm for cooling without additional mechanical cooling. Fault condition nine equation as defined by ASHRAE:"""]
    ],
    columns = ['fault_num', 'fault_cols', 'fault_def']
)

SENSORS_DF = pd.DataFrame(
    [
        ['mat', 'mixing air temperature', 'Mix Temp', 'temperature'],
        ['rat', 'return air temperature', 'Return Temp', 'temperature'],
        ['oat', 'outside air temperature', 'Out Temp', 'temperature'],
        ['economizer_sig', 'outside air damper position', 'AHU Dpr Cmd', 'percent'],
        ['clg', 'AHU cooling valve', 'AHU cool valv', 'percent'],
        ['htg', 'AHU heating valve', 'AHU htg valv', 'percent'],
        ['vav_total_air_flow', 'VAV total air flow', 'volume']
    ],
    columns = ['col_name', 'sensor_name', 'short_name', 'measurement type']
)


# parts of document:
#   - fault definition
#   - dataset plot
#   - dataset statistics
#   - summary stats and hist plot

# document generator:
    # makes plots
    # collates data with relevant text
    # organizes it all on the page
    #
    # it should not do any calculations!


# report calculator:
    # calculates all the data from corresponding faults
    # collates summary stats
    # passes them off to 

class DocumentGenerator:
    """Class provides the skeleton for creating a report document."""

    def __init__(self, fault_num, df):
        self.fault_num = fault_num
        self.document = Document()
        self.fault_cols = FAULT_COLS_DICT[fault_num]
        self.df = df
        self.output_col = f'fc{self.fault_num}_flag'

    # creates combination timeseries and fault flag plots 
    def create_dataset_plot(self) -> plt:
        df = self.df

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(25, 8))
        plt.title(f'Fault Conditions {self.fault_num} Plot')

        for col in self.fault_cols:
            col_short_name = SENSORS_DF[SENSORS_DF['col_name'] == col].iloc[0]['short_name']
            ax1.plot(df.index, df[col], label=col_short_name)

        ax1.legend(loc='best')
        ax1.set_ylabel('AHU Temps °F')

        ax2.plot(df.index, df[f'fc{self.fault_num}_flag'], label="Fault", color="k")
        ax2.set_xlabel('Date')
        ax2.set_ylabel('Fault Flags')
        ax2.legend(loc='best')

        plt.legend()
        plt.tight_layout()

        return fig

    # creates a histogram plot for times when fault condition is true
    def create_hist_plot(self) -> plt:
        df = self.df

        # calculate dataset statistics
        df[f"hour_of_the_day_fc{self.fault_num}"] = df.index.hour.where(df[self.output_col] == 1)

        # make hist plots fc10
        fig, ax = plt.subplots(tight_layout=True, figsize=(25, 8))
        ax.hist(df[f"hour_of_the_day_fc{self.fault_num}"].dropna())
        ax.set_xlabel("24 Hour Number in Day")
        ax.set_ylabel("Frequency")
        ax.set_title(f"Hour-Of-Day When Fault Flag {self.fault_num} is TRUE")
        return fig

    # give fault condition plain text definition/description and calculation
    def fault_def(self):
        self.document.add_heading(f"Fault Condition {self.fault_num} Report", 0)

        fault_defs = {
            1 : """Fault condition one of ASHRAE Guideline 36 is related to flagging poor performance of a AHU variable supply fan attempting to control to a duct pressure setpoint. Fault condition equation as defined by ASHRAE:""",
            2: """Fault condition two and three of ASHRAE Guideline 36 is related to flagging mixing air temperatures of the AHU that are out of acceptable ranges. Fault condition 2 flags mixing air temperatures that are too low and fault condition 3 flags mixing temperatures that are too high when in comparision to return and outside air data. The mixing air temperatures in theory should always be in between the return and outside air temperatures ranges. Fault condition two equation as defined by ASHRAE:""",
            9 : """Fault condition nine of ASHRAE Guideline 36 is an AHU economizer free cooling mode only with an attempt at flagging conditions where the outside air temperature is too warm for cooling without additional mechanical cooling. Fault condition nine equation as defined by ASHRAE:"""
        }

        p = self.document.add_paragraph(fault_defs[self.fault_num])

        self.document.add_picture(
            os.path.join(os.path.curdir, "images", f"fc{self.fault_num}_definition.png"),
            width=Inches(6),
        )

    def create_report(self,path: str) -> None:
        output_col = f"fc{self.fault_num}_flag"

        print(f"Starting {path} docx report!")

        self.fault_def()

        self.document.add_heading("Dataset Plot", level=2)

        fig = self.create_dataset_plot()
        fan_plot_image = BytesIO()
        fig.savefig(fan_plot_image, format="png")
        fan_plot_image.seek(0)

        # ADD IN SUBPLOTS SECTION
        self.document.add_picture(
            fan_plot_image,
            width=Inches(6),
        )


        self.document.add_heading("Dataset Statistics", level=2)

        calculator = ReportCalculator(self.fault_num, self.df)

        (
            total_days,
            total_hours,
            hours_fault_mode,
            percent_in_fault_mode,
            avg_fault_cond_vals,
            hours_motor_runtime,
            df_motor_on_filtered

        ) = calculator.summarize_fault_times()

        stats_lines = [
            f"Total time in days calculated in dataset: {total_days}",
            f"Total time in hours calculated in dataset: {total_hours}",
            f"Total time in hours for when fault flag is True: {hours_fault_mode}",
            f"Percent of time in the dataset when the fault flag is True: {percent_in_fault_mode}%",
            f"Percent of time in the dataset when the fault flag is False: {round((100 - percent_in_fault_mode), 2)}%",
            f"Calculated motor runtime in hours based off of VFD signal > zero: {hours_motor_runtime}"
        ]

        for line in stats_lines:
            paragraph = self.document.add_paragraph()
            paragraph.style = "List Bullet"
            paragraph.add_run(line)


        paragraph = self.document.add_paragraph()

        # if there is no faults skip the histogram plot
        fc_max_faults_found = self.df[output_col].max()
        if fc_max_faults_found != 0:

            # ADD HIST Plots
            self.document.add_heading("Time-of-day Histogram Plots", level=2)
            histogram_plot_image = BytesIO()
            histogram_plot = self.create_hist_plot()
            histogram_plot.savefig(histogram_plot_image, format="png")
            histogram_plot_image.seek(0)
            self.document.add_picture(
                histogram_plot_image,
                width=Inches(6),
            )

            paragraph = self.document.add_paragraph()

            max_faults_line = f'When fault condition {self.fault_num} is True, the'

            for (col_name,avg_val) in avg_fault_cond_vals.items():
                sensor_name = SENSORS_DF[SENSORS_DF['col_name'] == col_name].iloc[0]['sensor_name']
                max_faults_line =  f'{max_faults_line} average {sensor_name} is {avg_val} in °F, ' # need to add {units}

            max_faults_line = max_faults_line[0:-2] + "."

        else:
            print("NO FAULTS FOUND - For report skipping time-of-day Histogram plot")
            max_faults_line = f'No faults were found in this given dataset for the equation defined by ASHRAE.'


        paragraph.style = 'List Bullet'
        paragraph.add_run(max_faults_line)

        # ADD in Summary Statistics
        self.document.add_heading(
            'Summary Statistics filtered for when the AHU is running', level=1)

        # summ_stats_cols = [satsp_col, oat_col]

        for col in self.fault_cols:
            col_short_name = SENSORS_DF[SENSORS_DF['col_name'] == col].iloc[0]['short_name']
            # ADD in Summary Statistics
            self.document.add_heading(col_short_name, level=3)
            paragraph = self.document.add_paragraph()
            paragraph.style = 'List Bullet'
            paragraph.add_run(str(df_motor_on_filtered[col].describe()))

        self.document.add_heading("Suggestions based on data analysis", level=3)
        paragraph = self.document.add_paragraph()
        paragraph.style = "List Bullet"

        if percent_in_fault_mode > 5.0:
            paragraph.add_run(
                'The percent True metric that represents the amount of time for when the fault flag is True is high indicating temperature sensor error or the cooling valve is stuck open or leaking causing overcooling. Trouble shoot a leaking valve by isolating the coil with manual shutoff valves and verify a change in AHU discharge air temperature with the AHU running.')

        else:
            paragraph.add_run(
                'The percent True metric that represents the amount of time for when the fault flag is True is low inidicating the AHU components are within calibration for this fault equation Ok.')

        paragraph = self.document.add_paragraph()
        run = paragraph.add_run(f"Report generated: {time.ctime()}")
        run.style = "Emphasis"
        return self.document


# this_df = pd.read_csv("../ahu_data/hvac_random_fake_data/fc2_3_fake_data3.csv")
this_df = pd.read_csv("ahu_data/hvac_random_fake_data/fc2_3_fake_data1.csv", index_col="Date", parse_dates=True).rolling("5T").mean()

OUTDOOR_DEGF_ERR_THRES = 5.
MIX_DEGF_ERR_THRES = 5.
RETURN_DEGF_ERR_THRES = 2.

from faults import FaultConditionTwo
_fc2 = FaultConditionTwo(
    OUTDOOR_DEGF_ERR_THRES,
    MIX_DEGF_ERR_THRES,
    RETURN_DEGF_ERR_THRES,
    "mat",
    "rat",
    "oat",
    "supply_vfd_speed"
)

df2 = _fc2.apply(this_df)

# summarize_fault_times(this_df, 2, ["mixing_air_temperature", "return_air_temperature", "outside_air_temperature"])
# summarize_fault_times(this_df, 2, ["mat", "rat", "oat"])

report = ReportCalculator(2, this_df)
# print(report.summarize_fault_times())

# print(this_df)

document_generator = DocumentGenerator(2,this_df)

# fig = document_generator.create_dataset_plot()
# plt.show()

doc = document_generator.create_report('blah')


# path = os.path.join(os.path.curdir, "final_report")

# if not os.path.exists(path):
#     os.makedirs(path)
doc.save("test_gen.docx")
