# Gym Facilitation Timetable

The goal is to create a google app engine project that calls the google sheets API every month to create next month's schedule.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites

* Google Cloud Platform Account with billing project enabled
* Google CLI configured with google cloud platform 
* Google Sheets API enabled on project
* Created app engine account in IAM (Create a blank app engine project)
* App engine default service accoount given permission to access the relevant google docs / sheets

### Installing

Clone the repository

```
git clone https://github.com/orangecola/Gym-Facilitation-Timetable
```

Install third party libraries needed for use

```
pip install -t lib/ google-api-python-client
```

End with an example of getting some data out of the system or using it for a little demo

## Running the tests in dev_appserver.py
To test and develop your application in a development server

Create and download a .p12 key of your App Engine default service account from IAM & Admin > Service Accounts > Actions > Create Key > .p12
Convert the key to .pem format The code snippet below assumes the working directory is the location of the .p12 key. 
The .p12 key is called "secret.p12", the password for the key is "notasecret", and the output file is "secret.pem".

```
cat secret.p12 | openssl pkcs12 -nodes -nocerts -passin pass:notasecret | openssl rsa > secret.pem
```

Run the developmental app server with the service account created

```
dev_appserver.py --appidentity_email_address {EMAIL_ADDRESS} --appidentity_private_key_path {PATH_TO_KEY} {PATH_TO_CODE}
```

Access the development admin server from
```
http://localhost:8000
```

Remember that if login: admin is set, you will need to run the command from the cron menu instead of accessing it directly

## Deployment

Deplying to GCP

```
gcloud app deploy app.yaml
gcloud app deploy cron.yaml
```


## Built With

* [Google Sheets API](https://developers.google.com/sheets/api/)
* [Google App Engine](https://cloud.google.com/appengine/)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* https://medium.com/google-cloud/google-cloud-functions-scheduling-cron-5657c2ae5212
