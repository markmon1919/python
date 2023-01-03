__author__ = 'Mark Mon Monteros'


from pathlib import Path
import os, random, subprocess
import boto3
import pandas as pd
import pyperclip


class AWSProfileGenerator():

    def __init__(self):
        self.org = boto3.client('organizations')
        self.homedir = str(Path.home())
        self.accts = []
        self.acct = ''
        self.region = ''
        self.aws_cfg = self.homedir + '/.aws/config'
        self.ext_cfg = self.homedir + '/.aws/config-chrome-ext'

        self.acct = self.pull_accounts()
        self.get_region()
        self.add_aws_profile()
        self.add_ext_profile()

    def pull_accounts(self):
        print('\nPulling Latest Account List from Prod OU...')

        paginator = self.org.get_paginator('list_accounts_for_parent')
        page_iterator = paginator.paginate(ParentId='ou-mpfo-uv6625zp')
        for page in page_iterator:
            for acct in page['Accounts']:
                self.accts.append(acct)

        for i in self.accts:

            ps = subprocess.Popen(['grep', '-irnw', self.aws_cfg, '-e', i['Id']], 
                            stdout=subprocess.PIPE,
                            universal_newlines=True)
            output = ps.stdout.readline()
            return_code = ps.poll()
            
            if return_code == 1:
                self.acct = i

        if self.acct == '':
                print('\nNo New Accounts. Exiting...')
                self.exit()
        else:
            print('\nNew Account Detected :')
            print(u"\n\t\u2713 " + self.acct['Name'] + ' | ' + self.acct['Id'] + ' | ' + self.acct['Email'])
            return self.acct

    def get_region(self):
        # Download this worksheet --> https://sapphire365.sharepoint.com/sites/SapphireBeyondProgramme/_layouts/15/Doc.aspx?sourcedoc=%7B1417A7B3-CD5E-4ABF-A484-10F5E496496F%7D&file=Sapphire%20BeyondProgramme_Closure.pptx&action=edit&mobileredirect=true&CT=1671183012457&OR=ItemsView
        # Commented codes here are xlsx to csv conversion
        # xls_file = self.homedir + '/Downloads/Sapphire Migration Inventory.xlsx'
        # df = pd.read_excel(xls_file, 'Customer Contacts and Plan', index_col=None)
        # warnings.simplefilter(action='ignore', category=UserWarning)
        # df.to_csv(xls_file)
        # df.to_csv(xls_file.replace(xls_file[xls_file.rindex('.'):], '') + '.csv', sep=',', encoding='UTF-8')

        csv_file = self.homedir + '/Downloads/Sapphire Migration Inventory.csv'
        with open (csv_file, 'r', encoding='UTF-8') as accts_csv:
            df = pd.read_csv(accts_csv)
            df2 = df[['Name', 'AWS MBN Prod Account ', 'Region']]

        for data in df2.values:
            if self.acct['Email'] == data[1]:
                if data[2] == 'UK':
                    self.region = 'eu-west-2'
                elif data[2] == 'US':
                    self.region = 'us-east-2'
                else:
                    print('\nNo Region Data found in CSV File. Please specify region.')
                    self.region = self.input_region()
                break
        if self.region == '':
                print('\nNo Account Data found in CSV File. Please specify region.')
                self.region = self.input_region()

        print(u"\n\t\u2713 Region: " + self.region)

    def input_region(self):
        print('\nChoose Region:')
        print('\t[1] - eu-west-2 (London)')
        print('\t[2] - us-east-2 (Ohio)')
        print('\t[0] - Specify region name')
        print('\t[Any other keys] - default to eu-west-2 (London)')
        num = input('\nEnter number: ')
        match num:
            case "0":
                region = input('\nEnter region name: ')
                return region
            case "1":
                region = 'eu-west-2'
                return region
            case "2":
                region = 'us-east-2'
                return region
            case _:
                region = 'eu-west-2'
                return region

    def add_aws_profile(self):
        with open (self.aws_cfg, 'a') as aws_config:
            aws_config.write('\n[{profilename}]'.format(profilename=self.acct['Name']).lower())
            aws_config.write('\nrole_arn = arn:aws:iam::{accountId}:role/OrganizationAccountAccessRole'.format(accountId=self.acct['Id']))
            aws_config.write('\nregion = {region}'.format(region=self.region))
            aws_config.write('\nsource_profile = sapphire-payer\r')

        print('\nNew Profile added to ' + self.aws_cfg + '.\n')

    def add_ext_profile(self):
        with open (self.ext_cfg, 'a') as ext_config:
            ext_config.write('\n[{profilename}]'.format(profilename=self.acct['Name']).lower())
            ext_config.write('\nrole_arn = arn:aws:iam::{accountId}:role/OrganizationAccountAccessRole'.format(accountId=self.acct['Id']))
            ext_config.write('\nregion = {region}'.format(region=self.region))
            ext_config.write('\ncolor = {color}\r'.format(color=self.gen_hexcolor()))

        print('\nNew Profile added to ' + self.ext_cfg + '.\n')

        ps = subprocess.run(['cat', self.ext_cfg], 
                                    check=True, 
                                    capture_output=True)
        psNames = subprocess.run(['tail', '-n4'],
                                    input=ps.stdout,
                                    capture_output=True)
        output = psNames.stdout.decode('utf-8').strip()

        pyperclip.copy(output)

        print('\nNOTE: Paste clipboard now to AWS Extend Switch Roles Chrome Extension...\n')

    def gen_hexcolor(self):
        r = lambda: random.randint(0,255)
        return str('%02X%02X%02X' % (r(),r(),r())).lower()

    def exit(self):
        os._exit(0)


if __name__ == '__main__':
    print('\n>> AWS Profile Creator <<')
    print('\nCreated by: ' + __author__)
    
    print("\nNOTE: awsume to sapphire-payer profile before running script.\n")
    
    AWSProfileGenerator()
    
    print('\n\nDONE...!!!\n')
