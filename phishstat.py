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

import argparse
import database
import logging
import re
import statistic
import sys
import time
import xlsxwriter

from urlparse import urlparse, parse_qs

# Edit with the correct MySQL credentials
MYSQL_USER = "root"
MYSQL_PASSWORD = "toor"


def parse_args():
    """ Parse and validate the command line
    """
    parser = argparse.ArgumentParser(
        description=(
            "Automating the generation of statistics obtained from "
            "phishing engagements."
        )
    )

    parser.add_argument(
        "-c",
        "--client",
        dest="client",
        help="name for the database storing the engagement results",
        required=True,
        type=str
    )

    parser.add_argument(
        "-l",
        "--log",
        dest="apache_log",
        help="Apache log file",
        required=True,
        type=argparse.FileType('r')
    )

    parser.add_argument(
        "-p",
        "--parameter",
        dest="parameter",
        help="GET parametereter used to identify each clicks",
        required=True,
        type=str
    )

    parser.add_argument(
        "-r",
        "--ref",
        dest="email2ref",
        help="file containing the email to reference",
        required=True,
        type=argparse.FileType('r')
    )

    parser.add_argument(
        "-t",
        "--target",
        dest="target",
        help="file containg the target details",
        required=True,
        type=argparse.FileType('r')
    )

    parser.add_argument(
        "-v",
        "--verbose",
        action="store_const",
        const=logging.DEBUG,
        default=logging.INFO,
        dest="loglevel",
        help="enable verbose mode",
        required=False
    )

    return parser.parse_args()


def parse_apache_log(file, parameter):
    """ Return data from the Apache log
    """
    results = []
    # Extended Common Log Format
    pattern = re.compile(
        '([^ ]*) ([^ ]*) ([^ ]*) \[([^]]*)\] "([^"]*)" ([^ ]*) ([^ ]*)'
        ' "([^"]*)" "([^"]*)"'
    )

    for line in file:
            match = pattern.match(line)

            if match:
                (host, ignore, user, date, request,
                 status, size, referer, agent) = match.groups()

                if status == "200" and parameter in request:
                    parsed_url = urlparse(request.split()[1])

                    if parameter in parse_qs(parsed_url.query):
                        ref = ''.join(parse_qs(parsed_url.query)[parameter])
                        entries = ["link", host, agent, date, ref]
                        results.append(entries)
                elif (status == "200" and "download" in request and
                      parameter in referer):
                    parsed_url = urlparse(referer)

                    if parameter in parse_qs(parsed_url.query):
                        ref = ''.join(parse_qs(parsed_url.query)[parameter])
                        entries = ["download", host, agent, date, ref]
                        results.append(entries)
            else:
                logging.warning("could not parse: {}".format(line))
                pass

    return results


def main():
    try:
        args = parse_args()

        logging.basicConfig(
            format="%(levelname)-8s %(message)s",
            handlers=[
                logging.StreamHandler(sys.stdout)
            ],
            level=args.loglevel
        )

        # Variables summary
        logging.info("database name: {}".format(args.client))
        logging.info("target file: {}".format(args.target.name))
        logging.info("email to reference file: {}".format(args.email2ref.name))
        logging.info("Apache log: {}".format(args.apache_log.name))
        logging.info("GET parameter identifying clicks: {}".format(
            args.parameter
        ))

        # Setting-up and populating the database
        logging.info("initialising the database...")
        conn = database.init_db(args.client, MYSQL_USER, MYSQL_PASSWORD)

        logging.info("populating the 'Employee' table...")
        database.populate_Employee(conn, args.client, args.target)

        logging.info("populating the 'Link' table...")
        database.populate_Link(conn, args.email2ref)

        logging.info("parsing the Apache log...")
        results = parse_apache_log(args.apache_log, args.parameter)
        for entries in results:
            logging.debug("{}" .format(entries))

        logging.info("populating the 'Payload' table")
        database.populate_Payload(conn, results)

        # Generating stats
        workbook = xlsxwriter.Workbook(
            "phishstat-results_{}.xlsx".format(time.strftime("%Y%m%d-%H%M%S"))
        )

        logging.info("generating 'Click Types Per User' worksheet...")
        statistic.generate_click_types_per_user(conn, workbook)

        logging.info("generating 'Clicks Over Time' worksheet...")
        statistic.generate_clicks_over_time(conn, workbook)

        logging.info("generating 'Email Distribution' worksheet...")
        statistic.generate_email_distribution(conn, workbook)

        logging.info("generating 'Clicks Per Payload Type' worksheet...")
        statistic.generate_clicks_per_payload_type(conn, workbook)

        logging.info("generating 'Click Originating IP' worksheet...")
        statistic.generate_click_originating_ip(conn, workbook)

        logging.info("generating 'Users Who Did/Did Not Click' worksheet...")
        statistic.generate_users_who_did_and_did_not_click(conn, workbook)

        logging.info("generating 'Clicks Per Department'")
        statistic.generate_clicks_per_department(conn, workbook)

        workbook.close()
        conn.close
    except KeyboardInterrupt:
        logging.error("'CTRL+C' pressed, exiting...")


if __name__ == "__main__":
    main()
