class FaultCalculator:
    def __init__(self, fault_num, **kwargs):

        fault_cols_dict = {
            1: ['duct_static', 'supply_vfd_speed', 'duct_static_setpoint'],
            2: ['mat', 'rat', 'oat', 'supply_vfd_speed'],
            9: ['satsp', 'oat', 'economizer_sig', 'cooling_sig', 'supply_vfd_speed']
        }

        self.fault_num = fault_num
        self.fault_cols = fault_col_dicts[fault_num]

    def summarize_fault_times(self, df: pd.DataFrame, output_col: str = None) -> str:
        # if output_col is None:
        output_col = f"fc{fault_num}_flag"

        delta = df.index.to_series().diff()
        total_days = round(delta.sum() / pd.Timedelta(days=1), 2)

        total_hours = delta.sum() / pd.Timedelta(hours=1)

        hours_fault_mode = (delta * df[output_col]).sum() / pd.Timedelta(hours=1)

        percent_true = round(df[output_col].mean() * 100, 2)
        percent_false = round((100 - percent_true), 2)

        for col in input_cols:
            round(df[col].where(df(output_col)) == 1.mean(), 2)

        motor_on = df[self.fan_vfd_speed_col].gt(1.).astype(int)
        hours_motor_runtime = round(
            (delta * motor_on).sum() / pd.Timedelta(hours=1), 2)

        # for summary stats on I/O data to make useful
        df_motor_on_filtered = df[df[self.fan_vfd_speed_col] > 1.0]
        
        return (
            total_days,
            total_hours,
            hours_fault_mode,
            percent_true,
            percent_false,
            flag_true_oat,
            flag_true_satsp,
            hours_motor_runtime,
            df_motor_on_filtered
        )








class Report:
    """Class provides the definitions for a basic Fault Report."""

    def __init__(self, fault_num):
        self.fault_num = fault_num
        self.document = Document()


    def plot_fault_flag(self, df: pd.DataFrame, ax: plt.ax) -> ax:
        ax.plot(df.index, df[f'fc{fault_num}_flag'], label="Fault", color="k")
        ax.set_xlabel('Date')
        ax.set_ylabel('Fault Flags')
        ax.legend(loc='best')

        return ax

    def create_timeseries_plot(self, df: pd.DataFrame, output_col: str = None) -> plt:

        if output_col is None:
            output_col = f'fc{self.fault_num}_flag'

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(25, 8))
        plt.title(f'Fault Conditions {self.fault_num} Plot')

        ts_trace_cols

        for col in ts_trace_cols:
            ax1.plot(df.index, df[col], label = col.upper)

        # plot1a, = ax1.plot(df.index, df[self.satsp_col], label="SATSP")
        # plot1b, = ax1.plot(df.index, df[self.oat_col], label="OAT")
        ax1.legend(loc='best')
        ax1.set_ylabel('AHU Temps °F')

        # ax2.plot(df.index, df[f'fc{fault_num}_flag'], label="Fault", color="k")
        # ax2.set_xlabel('Date')
        # ax2.set_ylabel('Fault Flags')
        # ax2.legend(loc='best')

        ax2 = plot_fault_flag(self, df, ax2)

        plt.legend()
        plt.tight_layout()

        return fig

    def create_hist_plot(
        self, df: pd.DataFrame,
        output_col: str = None,
        vav_total_flow: str = None
    ) -> plt:

        if output_col is None:
            output_col = "fc9_flag"

        # calculate dataset statistics
        df["hour_of_the_day_fc9"] = df.index.hour.where(df[output_col] == 1)

        # make hist plots fc10
        fig, ax = plt.subplots(tight_layout=True, figsize=(25, 8))
        ax.hist(df.hour_of_the_day_fc9.dropna())
        ax.set_xlabel("24 Hour Number in Day")
        ax.set_ylabel("Frequency")
        ax.set_title(f"Hour-Of-Day When Fault Flag 9 is TRUE")
        return fig


    def fault_def(self):
        self.document.add_heading(f"Fault Condition {self.fault_num} Report", 0)

        fault_defs = {
            1 : """Fault condition one of ASHRAE Guideline 36 is related to flagging poor performance of a AHU variable supply fan attempting to control to a duct pressure setpoint. Fault condition equation as defined by ASHRAE:""",
            9 : """Fault condition nine of ASHRAE Guideline 36 is an AHU economizer free cooling mode only with an attempt at flagging conditions where the outside air temperature is too warm for cooling without additional mechanical cooling. Fault condition nine equation as defined by ASHRAE:"""
        }

        p = self.document.add_paragraph(fault_defs[self.fault_num])

        self.document.add_picture(
            os.path.join(os.path.curdir, "images", f"fc{fault_num}_definition.png"),
            width=Inches(6),
        )

        self.document.add_heading("Dataset Plot", level=2)



    def create_report(
        self,
        path: str,
        df: pd.DataFrame,
        output_col: str = None
    ) -> None:

        if output_col is None:
            output_col = f"fc{fault_num}_flag"

        print(f"Starting {path} docx report!")
        document = Document()
        document.add_heading(f"Fault Condition {self.fault_num} Report", 0)

        fault_defs = {
            1 : """Fault condition one of ASHRAE Guideline 36 is related to flagging poor performance of a AHU variable supply fan attempting to control to a duct pressure setpoint. Fault condition equation as defined by ASHRAE:""",
            9 : """Fault condition nine of ASHRAE Guideline 36 is an AHU economizer free cooling mode only with an attempt at flagging conditions where the outside air temperature is too warm for cooling without additional mechanical cooling. Fault condition nine equation as defined by ASHRAE:"""
        }

        p = document.add_paragraph(fault_defs[self.fault_num])

        document.add_picture(
            os.path.join(os.path.curdir, "images", f"fc{fault_num}_definition.png"),
            width=Inches(6),
        )

        document.add_heading("Dataset Plot", level=2)

        fig = self.create_timeseries_plot(df, output_col=output_col)
        fan_plot_image = BytesIO()
        fig.savefig(fan_plot_image, format="png")
        fan_plot_image.seek(0)

        # ADD IN SUBPLOTS SECTION
        document.add_picture(
            fan_plot_image,
            width=Inches(6),
        )

        document.add_heading("Dataset Statistics", level=2)

        (
            total_days,
            total_hours,
            hours_fault_mode,
            percent_true,
            percent_false,
            flag_true_oat,
            flag_true_satsp,
            hours_motor_runtime,
            df_motor_on_filtered

        ) = self.summarize_fault_times(df, output_col=output_col)

        stats_lines = [
            f"Total time in days calculated in dataset: {total_days}",
            f"Total time in hours calculated in dataset: {total_hours}",
            f"Total time in hours for when fault flag is True: {hours_fault_mode}",
            f"Percent of time in the dataset when the fault flag is True: {percent_true}%",
            f"Percent of time in the dataset when the fault flag is False: {percent_false}%",
            f"Calculated motor runtime in hours based off of VFD signal > zero: {hours_motor_runtime}"
        ]

        for line in stats_lines:
            paragraph = document.add_paragraph()
            paragraph.style = "List Bullet"
            paragraph.add_run(line)


        paragraph = document.add_paragraph()

        # if there is no faults skip the histogram plot
        fc_max_faults_found = df[output_col].max()
        if fc_max_faults_found != 0:

            # ADD HIST Plots
            document.add_heading("Time-of-day Histogram Plots", level=2)
            histogram_plot_image = BytesIO()
            histogram_plot = self.create_hist_plot(df, output_col=output_col)
            histogram_plot.savefig(histogram_plot_image, format="png")
            histogram_plot_image.seek(0)
            document.add_picture(
                histogram_plot_image,
                width=Inches(6),
            )

            paragraph = document.add_paragraph()

            max_faults_line = f'When fault condition 9 is True the average outside air is {flag_true_oat} in °F and the supply air temperature setpoinht is {flag_true_satsp} in °F.'

        else:
            print("NO FAULTS FOUND - For report skipping time-of-day Histogram plot")
            max_faults_line = f'No faults were found in this given dataset for the equation defined by ASHRAE.'


        paragraph.style = 'List Bullet'
        paragraph.add_run(max_faults_line)

        # ADD in Summary Statistics
        document.add_heading(
            'Summary Statistics filtered for when the AHU is running', level=1)

        summ_stats_cols = [satsp_col, oat_col]

        for col in summ_stats_cols:
            # ADD in Summary Statistics
            document.add_heading(col.upper, level=3)
            paragraph = document.add_paragraph()
            paragraph.style = 'List Bullet'
            paragraph.add_run(str(df_motor_on_filtered[col].describe()))

        document.add_heading("Suggestions based on data analysis", level=3)
        paragraph = document.add_paragraph()
        paragraph.style = "List Bullet"

        if percent_true > 5.0:
            paragraph.add_run(
                'The percent True metric that represents the amount of time for when the fault flag is True is high indicating temperature sensor error or the cooling valve is stuck open or leaking causing overcooling. Trouble shoot a leaking valve by isolating the coil with manual shutoff valves and verify a change in AHU discharge air temperature with the AHU running.')

        else:
            paragraph.add_run(
                'The percent True metric that represents the amount of time for when the fault flag is True is low inidicating the AHU components are within calibration for this fault equation Ok.')

        paragraph = document.add_paragraph()
        run = paragraph.add_run(f"Report generated: {time.ctime()}")
        run.style = "Emphasis"
        return document

