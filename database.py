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

from datetime import datetime, timedelta


def init_db(name, username, password):
    """ Create a database specific to the engagement
    """
    statement = (
        """
        CREATE DATABASE {0} DEFAULT CHARACTER SET 'utf8';
        use {0};

        CREATE TABLE Employee (
            client          VARCHAR(32)     NOT NULL,
            email           VARCHAR(255)    NOT NULL,
            role            VARCHAR(16),
            PRIMARY KEY(email, client)
        ) ENGINE InnoDB DEFAULT CHARSET=utf8;

        CREATE TABLE Link (
            ref             VARCHAR(16)     NOT NULL,
            email           VARCHAR(255)    NOT NULL,
            PRIMARY KEY(ref),
            FOREIGN KEY(email) REFERENCES Employee(email)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        ) ENGINE InnoDB DEFAULT CHARSET=utf8;

        -- This table store information such as login and password
        -- Customise it according to the needs of the engagement
        CREATE TABLE Payload (
            idP             INTEGER         NOT NULL AUTO_INCREMENT,
            type            VARCHAR(16)     NOT NULL,
            host            VARCHAR(16)     NOT NULL,
            agent           VARCHAR(256),
            date            TIMESTAMP       NOT NULL,
            ref             VARCHAR(16)     NOT NULL,
            PRIMARY KEY(idP),
            FOREIGN KEY(ref) REFERENCES Link(ref)
                ON DELETE CASCADE
                ON UPDATE CASCADE
        ) ENGINE InnoDB DEFAULT CHARSET=utf8;
        """.format(name)
    )

    try:
        conn = MySQLdb.connect(user=username, passwd=password)
        cursor = conn.cursor()
        cursor.execute(statement)

        cursor.close()

        return conn
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def populate_Employee(conn, client, file):
    """ Populate the 'Employee' table from the target file
    """
    try:
        cursor = conn.cursor()

        for line in file:
            email = line.split(',')[0]
            role = line.split(',')[1].rstrip()
            statement = (
                "INSERT INTO Employee(client, email, role) "
                "VALUES(\"{}\", \"{}\", \"{}\");".format(client, email, role)
            )

            logging.debug("{}".format(statement))
            cursor.execute(statement)

        conn.commit()
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def populate_Link(conn, file):
    """ Populate the 'Link' table from the email to ref file
    """
    try:
        cursor = conn.cursor()

        for line in file:
            email = line.split(',')[0]
            ref = line.split(',')[1].rstrip()
            statement = (
                "INSERT INTO Link(ref, email) VALUES(\"{}\", \"{}\");".format(
                    ref, email
                )
            )

            logging.debug("{}".format(statement))
            cursor.execute(statement)

        conn.commit()
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)


def populate_Payload(conn, results):
    """ Populate the 'Payload' table from the Apache log parsing results
    """
    try:
        cursor = conn.cursor()

        for entries in results:
            date = datetime.strptime(
                entries[3].split()[0],
                "%d/%b/%Y:%H:%M:%S"
            )
            statement = (
                "INSERT INTO Payload(type, host, agent, date, ref) VALUES "
                "(\"{}\", \"{}\", \"{}\", \"{}\", \"{}\");".format(
                    entries[0], entries[1], entries[2], date, entries[4]
                )
            )

            logging.debug("{}".format(statement))
            cursor.execute(statement)

        conn.commit()
        cursor.close()
    except (MySQLdb.Error) as ex:
        logging.error("{}".format(ex))
        sys.exit(1)
