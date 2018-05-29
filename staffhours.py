"""
    Entry Dates will be compared to the reporting period.  Entries that occured prior to the
    reporting period will have only the associated services created during the reporting month
    returned.  Otherwise, services up to three months prior to entry will be returnedself.
    All services will be returned as a total hours by case manager pivot table.
"""

__author__ = "David Marienburg"

import pandas as pd
import numpy as np

from datetime import datetime as dt
from tkinter.filedialog import askopenfilename as aofn
from tkinter.filedialog import asksaveasfilename as asfn

class StaffHours:
    def __init__(self, report, service_chart, period_start, period_end, eha):
        """
        :report: a file location for the Staff Hours.xlsx report
        :service_chart: a file location for the most recent Services Chart
        :period_start: a datetime.datetime date object
        :period_end: a datetime.datetime date object
        """

        self.report = report
        self.service_chart = service_chart
        self.period_start = period_start
        self.period_end = period_end
        self.eha = eha
        self.services_chart_df, self.eha_services_chart_df = self.create_services_chart(self.service_chart)
        self.entries, self.services = self.create_data_frames(self.report)

    def create_services_chart(self, chart):
        standard_chart = pd.read_excel(chart, sheet_name="Reporting Use")
        eha_chart = pd.read_excel(chart, sheet_name="EHA & EHA2 Reporting")
        return standard_chart[[
            "Service Provider Specific Code",
            "Time Value"
        ]], eha_chart[[
            "Service Provider Specific Code",
            "Time Value"
        ]]

    def create_data_frames(self, report):
        entry_data = pd.read_excel(report, sheet_name="Entry Data")
        service_data = pd.read_excel(report, sheet_name="Service Data")
        return entry_data[[
            "Client Uid",
            "Entry Exit Entry Date"
        ]], service_data[[
            "Client Uid",
            "Service Provide Start Date",
            "Service User Creating",
            "Provider Specific Code"
        ]]

    def check_entry_date(self, service_df, entry_df):
        # create merged version of the service_df and entry_df
        data = pd.merge(service_df, entry_df, on="Client Uid", how="outer")

        # strip the staff user id number out of the Service User Creating column
        data["Service User Creating"] = data["Service User Creating"].str.replace("\([\d]*\)","")

        # add the period_start and period_end variables to similarly named columns
        data["Period Start"] = self.period_start
        data["Period End"] = self.period_end

        # create a datetime column containing dates 3 months prior to the entry date
        data["Prior to Start"] = data["Entry Exit Entry Date"] - pd.DateOffset(months=3)

        # create a dataframe of services created during the reporting period for participants who
        # entered prior to the reporting period.
        start_in_period = data[
            (
                data["Entry Exit Entry Date"] < data["Period Start"]
            ) &
            (
                (data["Service Provide Start Date"] > data["Period Start"]) |
                (data["Service Provide Start Date"] == data["Period Start"])
            )
        ]

        # create a dataframe of services entered up to three months prior to the entry into the
        # provider for participants who entered the provider during the reporting period.
        start_prior_to_period = data[
            (
                (data["Entry Exit Entry Date"] == data["Period Start"]) |
                (data["Entry Exit Entry Date"] > data["Period Start"])
            ) &
            (
                (data["Service Provide Start Date"] == data["Prior to Start"]) |
                (data["Service Provide Start Date"] > data["Prior to Start"])
            )
        ]

        # merge the two dataframes
        merge_list = [start_in_period, start_prior_to_period]
        merged = pd.concat(merge_list, join="outer", ignore_index=True)

        # return the merged results stripping out extraneous columns
        return merged[[
            "Client Uid",
            "Service Provide Start Date",
            "Service User Creating",
            "Provider Specific Code"
        ]]

    def find_service_hours(self, relevant_services, eha=False):
        if eha:
            hours = pd.merge(
                relevant_services,
                self.eha_services_chart_df,
                left_on="Provider Specific Code",
                right_on="Service Provider Specific Code",
                how="left"
            )
        else:
            hours = pd.merge(
                relevant_services,
                self.services_chart_df,
                left_on="Provider Specific Code",
                right_on="Service Provider Specific Code",
                how="left"
            )
        return hours

    def create_pivot_table(self, hours_df):
        pivot = pd.pivot_table(
            hours_df,
            index=["Service User Creating"],
            values=["Time Value"],
            aggfunc={"Time Value": np.sum}
        )
        return pivot

    def save_report(self, processed_df, pivot_table):
        # initialize the writer object
        writer = pd.ExcelWriter(asfn(title="Save the Processed Report"), engine="xlsxwriter")

        # send dataframes to the writer object and assign a sheet_name
        pivot_table.to_excel(writer, sheet_name="Summary")
        processed_df.to_excel(writer, sheet_name="Processed Data", index=False)
        self.services.to_excel(writer, sheet_name="Raw Services", index=False)
        self.entries.to_excel(writer, sheet_name="Raw Entries", index=False)
        self.eha_services_chart_df.to_excel(writer, sheet_name="EHA Services Chart", index=False)
        self.services_chart_df.to_excel(writer, sheet_name="Services Chart", index=False)

        # complete the file saving
        writer.save()

    def process(self):
        data = self.check_entry_date(self.services, self.entries)
        hours_data = self.find_service_hours(data, self.eha)
        pivot = self.create_pivot_table(hours_data)
        self.save_report(hours_data, pivot)

if __name__ == "__main__":
    report = aofn(title="Open the Staff Hours Report")
    charts = aofn(title="Open the Service Chart")
    start = dt(year=2018, month=4, day=1)
    end = dt(year=2018, month=5, day=1)
    run = StaffHours(report, charts, start, end, eha=False)
    run.process()
