# Property Management

First Class Property Management is a provider of Temporary Accommodation to a number of London Councils.

The provider leases properties from landlords.

Landlords are paid monthly by First Class Property Management.

The provider agrees and gets paid a nightly rate for each client supplied by the London Councils.

Each booking includes the client name, phone number, and nightly rate.


# Architecture

A web page is used for entering data and to generate reports.

## Data

A database is used for storing and querying property management data.

### Reports

The following reports are provided:

- Landlord Monthly Payment Report (showing how much each landlord should be paid for the month based on bookings)
- Council Billing Report (what each council owes, based on clients placed)
- Current Occupancy Report (which clients are currently in which properties)
- Property Utilization Report (percentage of days each property was booked in the last 30 days)
- Landlord Payment History (all payments recorded)
- Council Client Nights Report (number of nights each council’s clients occupied properties in a period)

## Data input

There is a `Client Intake` Form.

# Implementation

# Infrastructure

AWS is used for the infrastructure and is automated using Terraform.

The following AWS services will be used:
- CloudFront, S3, Lambda for the web page (with Next.js and Node.js)
- MySQL for the managed database
