#!/usr/bin/env python
#    Copyright (C) 2016 - 2017 Alexandre Teyar

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

# http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
#    limitations under the License.

import logging
import MySQLdb
import sys
import xlsxwriter

from datetime import datetime, timedelta


def write_worksheet(workbook, worksheet, table_headers, table_data):
    """ Create an Excel worksheet containing the 'table_headers'
        and 'table_data' dataset
    """
    if not table_data:
        logging.warning("'{}' could not be created".format(
            worksheet
        ))
        return
    else:
        worksheet = workbook.add_worksheet("{}".format(worksheet))

        column_count = 0
        row_count = 0
        table_column_count = len(table_headers) - 1
        table_row_count = len(table_data)

        logging.debug("{}".format(table_headers))
        logging.debug("{}".format(table_data))

        worksheet.add_table(
            row_count,
            column_count,
            table_row_count,
            table_column_count,
            {
                "banded_rows": True,
                "columns": table_headers,
                "data": table_data,
                "first_column": True,
                "style": "Table Style Medium 1"
            }
        )


def generate_click_types_per_user(conn, workbook):
    """ Create an Excel worksheet containing the 'Click Types Per User' dataset
    """
    table_data = []
    table_headers = [
        {"header": "Reference"},
        {"header": "Download File"},
        {"header": "Clicked Link"},
        {"header": "Total"}
    ]

    try:
        cursor = conn.cursor()
        statement = (
            """
            SELECT DISTINCT
                Payload.ref
            FROM
                Payload
            ORDER BY
                Payload.ref
            """
        )

        cursor.execute(statement)
        results = cursor.fetchall()

        for result in results:
            ref = result[0]
            statement = (
                """
                SELECT
                    COUNT(CASE WHEN Payload.type = 'Download' THEN 1 END),
                    COUNT(CASE WHEN Payload.type = 'Link' THEN 1 END)
                FROM
                    Payload
                WHERE
                    Payload.ref = '{}';
                """.format(ref)
            )

            cursor.execute(statement)
            results = cursor.fetchone()
            download_count = results[0]
            link_count = results[1]
            total = download_count + link_count

            table_data.append([ref, download_count, link_count, total])

        write_worksheet(
            workbook, "Click Types Per User",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def generate_clicks_over_time(conn, workbook):
    """ Create an Excel worksheet containing the 'Clicks Over Time' dataset
    """
    table_data = []
    table_headers = [
        {"header": "Time"},
        {"header": "Click on Download"},
        {"header": "Click on Link"}
    ]

    try:
        cursor = conn.cursor()

        statement = (
            """
            SELECT
                MIN(Payload.date),
                MAX(Payload.date)
            FROM Payload;
            """
        )

        cursor.execute(statement)
        results = cursor.fetchone()
        min_date = results[0]
        max_date = results[1]

        date = min_date

        while date < max_date:
            statement = (
                """
                SELECT
                    COUNT(CASE WHEN Payload.type = 'Download' THEN 1 END) AS 'Download',
                    COUNT(CASE WHEN Payload.type = 'Link' THEN 1 END) AS 'Link'
                FROM
                    Payload
                WHERE
                    Payload.date <= '{}';
                """.format(date)
            )

            cursor.execute(statement)
            result = cursor.fetchone()
            download_count = result[0]
            link_count = result[1]

            logging.debug(
                "{} {} {}".format(
                    date.strftime("%H:%M:%S"), download_count, link_count
                )
            )

            table_data.append([
                date.strftime("%H:%M:%S"),
                download_count,
                link_count
            ])

            # Get stats for every 1 min of the engagement
            date = date + timedelta(minutes=1)

        write_worksheet(
            workbook, "Clicks Over Time",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def generate_email_distribution(conn, workbook):
    """ Create an Excel worksheet containing the 'Email Distribution' dataset
    """
    table_data = []
    table_headers = [
        {"header": "Department"},
        {"header": "# of Employee"}
    ]

    try:
        cursor = conn.cursor()

        statement = (
            """
            SELECT DISTINCT
                Employee.role
            FROM
                Employee
            ORDER BY
                Employee.role;
            """
        )

        cursor.execute(statement)
        results = cursor.fetchall()

        for result in results:
            department = result[0]
            statement = (
                """
                SELECT DISTINCT
                    COUNT(*)
                FROM
                    Employee
                WHERE
                    Employee.role = '{}';
                """.format(department)
            )

            cursor.execute(statement)
            result = cursor.fetchone()
            employee_count = result[0]

            table_data.append([department, employee_count])

        write_worksheet(
            workbook, "Email Distribution",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def generate_clicks_per_payload_type(conn, workbook):
    """ Create an Excel worksheet containing the 'Clicks Per Payload Type' dataset
    """
    table_data = []
    table_headers = [
        {"header": "Type"},
        {"header": "# of Click"}
    ]

    try:
        cursor = conn.cursor()

        statement = (
            """
            SELECT
                COUNT(CASE WHEN Payload.type = 'Download' THEN 1 END) AS 'Download',
                COUNT(CASE WHEN Payload.type = 'Link' THEN 1 END) AS 'Link'
            FROM
                Payload;
            """
        )

        cursor.execute(statement)
        result = cursor.fetchone()
        download_count = result[0]
        link_count = result[1]

        table_data.append(["Download", download_count])
        table_data.append(["Link", link_count])

        write_worksheet(
            workbook, "Clicks Per Payload Type",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def generate_click_originating_ip(conn, workbook):
    """ Create an Excel worksheet containing the 'Click Originating IP' dataset
    """
    table_data = []
    table_headers = [
        {"header": "Source IP"},
        {"header": "#"}
    ]

    try:
        cursor = conn.cursor()

        statement = (
            """
            SELECT DISTINCT
                Payload.host
            FROM
                Payload;
            """
        )

        cursor.execute(statement)
        results = cursor.fetchall()

        for result in results:
            host = result[0]
            statement = (
                """
                SELECT DISTINCT
                    COUNT(*)
                FROM
                    Payload
                WHERE
                    Payload.host = '{}';
                """.format(host)
            )

            cursor.execute(statement)
            result = cursor.fetchone()
            host_count = result[0]

            table_data.append([host, host_count])

        write_worksheet(
            workbook, "Click Originating IP",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def generate_users_who_did_and_did_not_click(conn, workbook):
    """ Create an Excel worksheet containing
        the 'User Who Did/Did Not Click' dataset
    """
    table_data = []
    table_headers = [
        {"header": " "},
        {"header": "Success Rate"}
    ]

    try:
        cursor = conn.cursor()

        statement = (
            """
            SELECT
                COUNT(Link.email)
            FROM
                Link
            WHERE
                EXISTS (SELECT 1 FROM Payload WHERE Link.ref = Payload.ref);
            """
        )

        cursor.execute(statement)
        result = cursor.fetchone()
        click_count = result[0]

        statement = (
            """
            SELECT
                COUNT(Link.email)
            FROM
                Link
            WHERE
                NOT EXISTS (SELECT 1 FROM Payload WHERE Link.ref = Payload.ref);
            """
        )

        cursor.execute(statement)
        result = cursor.fetchone()
        not_click_count = result[0]

        table_data.append(["Users Clicked", click_count])
        table_data.append(["Users Not Clicked", not_click_count])

        write_worksheet(
            workbook, "User Who Did (Not) Click",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


# To be Finished
def generate_clicks_per_department(conn, workbook):
    """ Create an Excel worksheet containing the 'Clicks Per Department' dataset
    """
    table_data = []
    table_headers = [
        {"header": "Department"},
        {"header": "Clicked on Download"},
        {"header": "Clicked on Email Link"},
        {"header": "Total"}
    ]

    try:
        cursor = conn.cursor()

        statement = (
            """
                SELECT DISTINCT
                    Employee.role
                FROM
                    Employee
                ORDER BY
                    Employee.role;
            """
        )

        cursor.execute(statement)
        results = cursor.fetchall()

        for result in results:
            department = result[0]
            statement = (
                """
                SELECT
                    COUNT(Payload.idP)
                FROM
                    Payload
                WHERE
                    Payload.ref
                IN (
                    SELECT
                        Link.ref
                    FROM
                        Link
                    WHERE
                        Link.email
                    IN (
                        SELECT
                            Employee.email
                        FROM
                            Employee
                        WHERE
                            Employee.role = '{}'
                    )
                )
                AND
                    Payload.type = 'Download';
                """.format(department)
            )

            cursor.execute(statement)
            result = cursor.fetchone()
            download_count = result[0]

            statement = (
                """
                SELECT
                    COUNT(Payload.idP)
                FROM
                    Payload
                WHERE
                    Payload.ref
                IN (
                    SELECT
                        Link.ref
                    FROM
                        Link
                    WHERE
                        Link.email
                    IN (
                        SELECT
                            Employee.email
                        FROM
                            Employee
                        WHERE
                            Employee.role = '{}'
                    )
                )
                AND
                    Payload.type = 'Link';
                """.format(department)
            )

            cursor.execute(statement)
            result = cursor.fetchone()
            link_count = result[0]
            total = download_count + link_count

            table_data.append([department, download_count, link_count, total])

        write_worksheet(
            workbook, "Clicks Per Department",
            table_headers, table_data
        )
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)
