#!/usr/bin/env python
# coding: utf-8

# Option Smile Volatility and Implied Probabilities Analysis
# 
# This project aims to analyze the implications of concavity in Implied Volatility (IV) curves. The primary focus is on the option smile volatility and the calculation of implied probabilities based on the observed market data.
# 
# Companies Analyzed:
# - BLK: Black Rock
# - GOOGL: Google A
# - JPM: JP Morgan and Chase
# - META: Meta (formerly Facebook)
# - FLX: Netflix
# - PEP: PEPSICO
# - WBA: Walgreens
# 
# Data Enhancements:
# The dataset is supplemented with the Dividend Yield for specific years and the risk-free rate to provide a holistic view of the market conditions and to facilitate a comprehensive analysis.
# 

# In[1]:


# Required Libraries and Packages

# pandas: A powerful data manipulation and analysis library for Python.
import pandas as pd

# numpy: A library for numerical computing in Python.
import numpy as np

# matplotlib.pyplot: A plotting library for creating visualizations in Python.
import matplotlib.pyplot as plt

# math: A standard Python library for mathematical operations.
import math as m

# os: A standard Python library for interacting with the operating system.
import os as os

# scipy.stats.norm: Functions for working with the normal distribution.
from scipy.stats import norm

# py_vollib.black_scholes_merton: A library for calculating Black-Scholes-Merton option prices.
from py_vollib.black_scholes_merton import black_scholes_merton as bs

# py_vollib.black_scholes_merton.greeks.analytical: Functions for calculating Greeks using the analytical method.
from py_vollib.black_scholes_merton.greeks.analytical import delta, gamma, vega, theta, rho

# xlsxwriter: A Python module for writing files in the Excel 2007+ XLSX file format.
import xlsxwriter 

# random: A standard Python library for generating random numbers.
import random

# datetime: A standard Python library for working with dates and times.
import datetime


# In[2]:


class InputOutput:
    """
    This class provides methods for reading from and writing to Excel files.
    
    Attributes:
    - _input_: The base directory where the data files are located.
    - company_name: The name of the company.
    - folder_name: The name of the folder under the company directory.
    - file_name: The name of the Excel file (without the .xlsx extension).
    - sheet_names: A list of sheet names in the Excel file.
    """
#    _input_ = '/Users/darshkachhara/Desktop/Project_fe800/Data/'
    def __init__(self, _input,company_name, folder_name, file_name,sheet_names):
        """
           
        Initializes the InputOutput with directory and file details.
        
        Parameters:
        - _input (str): Base directory path.
        - company_name (str): Name of the company.
        - folder_name (str): Name of the folder under the company directory.
        - file_name (str): Name of the Excel file (without the .xlsx extension).
        - sheet_names (list): List of sheet names in the Excel file.
        
        """
            
        self._input_=_input
        self.company_name = company_name
        self.folder_name = folder_name
        self.file_name = file_name
        self.sheet_names=sheet_names
    def excel_open(self,sheet_name):
        
        """
        Opens and reads the specified sheet from the Excel file.
        
        Parameters:
        - sheet_name (str): Name of the sheet to be read.
        
        Returns:
        - DataFrame: A pandas DataFrame containing the data from the specified sheet.
        """
        
        file=f"{self._input_}/{self.company_name}/{self.folder_name}/{self.file_name}"
        read_file=pd.read_excel(file,sheet_name)
        return read_file
        
    def __del__(self):
        """
        Destructor method.
        """
        return 0


# In[3]:


class Calculate:
    
    """
    The Calculate class is designed to perform various financial calculations related to options trading.
    It provides functionalities such as curve fitting for implied volatility, option pricing using the 
    Black-Scholes model, calculation of Greeks, and computation of implied probability.
    
    """
    
    #trade_off_lambda=0.01
    def __init__(self, file: pd.DataFrame):
        
        """
        Initializes the Calculate class with data from the provided file.
        
        Parameters:
        - file (DataFrame): A DataFrame containing important financial data such as 'implied_foward', 
        'strike', 'moneyness', 'implied_vol_val', 
        'days_to_expiry', 'current_price', 'implied_vol_percent', 
        'days_to_expiry_yearformat', 'Company_name', 'Q_name', 
        'Current_Date', 'Expiry_Number', 'Expiry_Date', 'Dividend', 'rf_rate','FN'
        
        Be mindful of the column names in the excel sheet as they are essential for the algo to run
        
        """
        # List of attributes to be fetched from the file
        attributes = [
        'implied_foward', 'strike', 'moneyness', 'implied_vol_val', 
        'days_to_expiry', 'current_price', 'implied_vol_percent', 
        'days_to_expiry_yearformat', 'Company_name', 'Q_name', 
        'Current_Date', 'Expiry_Number', 'Expiry_Date', 'Dividend', 'rf_rate',
        'FN']
        # Setting each attribute from the file to the class instance
        for attr in attributes:
            setattr(self, attr, file[attr])
        
    def curve_fitting_implied_vol(self,trade_off_lambda):
        
        """
        Performs curve fitting for implied volatility using a tridiagonal matrix approach. This method 
        aims to generate a smooth implied volatility curve from market data.
        
        Parameters:
        - trade_off_lambda (float): A control variable that determines the trade-off between the smoothness 
                                    of the curve and its fit to the market data.
        
        Returns:
        - pd.Series: A series containing the calculated implied volatilities for each moneyness level.
        """

         # Define the length of the implied volatility percent
        total_len = len(self.implied_vol_percent)

        # Initialize an empty matrix for rows
        list_of_rows = np.empty((0, total_len))

        # Count non-zero implied volatility percent values
        I = (self.implied_vol_percent != 0).sum()

        # Compute the difference matrix for moneyness values
        diff_matrix = np.abs(self.moneyness.values[:, None] - self.moneyness.values)

        # Calculate the minimum non-zero difference
        d = round(np.min(diff_matrix[diff_matrix > 0]), 3)

        # Compute the known sigma vector based on the trade-off lambda
        known_sigma_vector = (trade_off_lambda * total_len / (I * d**4)) * self.implied_vol_percent

        # Calculate the bracket term for the matrix
        bracket_term = 6 + (trade_off_lambda * total_len / (I * d**4))
        
        for i in range(0,total_len):
            row_of_matrix = [0]*len(self.implied_vol_percent)
            if(i==0): #special condition
                if(self.implied_vol_percent.iloc[i]==0):
                    row_of_matrix[0]=3
                    row_of_matrix[1]=-4
                    row_of_matrix[2]=1
                else:
                    row_of_matrix[0]=-3+bracket_term
                    row_of_matrix[1]=-4
                    row_of_matrix[2]=1
            elif(i==1):#special condition
                if(self.implied_vol_percent.iloc[i]==0):
                    row_of_matrix[0]=-3
                    row_of_matrix[1]=6
                    row_of_matrix[2]=-4
                    row_of_matrix[3]=1
                else:
                    row_of_matrix[0]=-3
                    row_of_matrix[1]=bracket_term
                    row_of_matrix[2]=-4
                    row_of_matrix[3]=1
            elif(i==total_len-2):#special condition
                if(self.implied_vol_percent.iloc[i]==0):
                    row_of_matrix[i-2]=1
                    row_of_matrix[i-1]=-4
                    row_of_matrix[i]=6
                    row_of_matrix[i+1]=-3
                else:
                    row_of_matrix[i-3]=1
                    row_of_matrix[i-2]=-4
                    row_of_matrix[i]=bracket_term
                    row_of_matrix[i+1]=-3
            elif(i==total_len-1):#special condition
                if(self.implied_vol_percent.iloc[i]==0):
                    row_of_matrix[i-2]=1
                    row_of_matrix[i-1]=-4
                    row_of_matrix[i]=3
                else:
                    row_of_matrix[i-2]=1
                    row_of_matrix[i-1]=-4
                    row_of_matrix[i]=bracket_term-3
            else:

                if(self.implied_vol_percent.iloc[i]==0):
                    row_of_matrix[i-2]=1
                    row_of_matrix[i-1]=-4
                    row_of_matrix[i]=6
                    row_of_matrix[i+1]=-4
                    row_of_matrix[i+2]=1
                else:
                    row_of_matrix[i-2]=1
                    row_of_matrix[i-1]=-4
                    row_of_matrix[i]=bracket_term
                    row_of_matrix[i+1]=-4
                    row_of_matrix[i+2]=1

            list_of_rows= np.append(list_of_rows, [row_of_matrix], axis=0)
        A_mat = np.matrix(list_of_rows) #create the matrix to solve AX=B matrix where B is the vector 
        #containig either observable vol values or 0

        sigma_vector = np.array(known_sigma_vector)#solves the equation for X
        unknown_sigma_vector = np.linalg.solve(A_mat,sigma_vector)
        return(pd.Series(unknown_sigma_vector))#return value
    
    def prices(self) -> pd.DataFrame:
        
        """
        Calculates the prices of call and put options using the Black-Scholes formula. This method uses 
        the class attributes to fetch the necessary parameters for the Black-Scholes formula.
        
        Returns:
        - pd.DataFrame: A dataframe containing the calculated call and put option prices for each strike.
        """
        
        call_prices = []
        put_prices = []
        for curr_price, strike, expiry, rf, iv, div in zip(
            self.current_price, self.strike, self.days_to_expiry_yearformat, 
            self.rf_rate, self.curve_fitting_implied_vol(0.01), self.Dividend):
            call_prices.append(bs('c', curr_price, strike, expiry, rf, iv, div))
            put_prices.append(bs('p', curr_price, strike, expiry, rf, iv, div))
        return pd.DataFrame({
            'Call_P': call_prices,
            'Put_P': put_prices
        })

    def Greeks(self) -> pd.DataFrame:
        
        """
        Calculates the Greeks (Delta, Gamma, Vega, Theta, Rho) for options. Greeks are essential measures 
        in options trading representing the sensitivity of the option price to various factors.
        
        Returns:
        - pd.DataFrame: A dataframe containing the calculated values of the Greeks for each option.
        """

        results = {
            'call_delta': [],
            'put_delta': [],
            'gamma': [],
            'vega': [],
            'theta_c': [],
            'theta_p': [],
            'rho_c': [],
            'rho_p': []
        }

        greek_functions = {
            'delta': delta,
            'gamma': gamma,
            'vega': vega,
            'theta': theta,
            'rho': rho
        }

        for curr_price, strike, expiry, rf, iv, div in zip(
                self.current_price, self.strike, self.days_to_expiry_yearformat, 
                self.rf_rate, self.curve_fitting_implied_vol(0.01), self.Dividend):

            # Calculate call and put delta
            call_d = delta('c', curr_price, strike, expiry, rf, iv, div)
            results['call_delta'].append(call_d)
            results['put_delta'].append(call_d - 1)

            # Calculate other Greeks
            for greek, func in greek_functions.items():
                if greek != 'delta':  # Delta is already calculated
                        if greek in ['theta', 'rho']:
                            for opt_type in ['c', 'p']:
                                key = f"{greek}_{opt_type}" 
                                results[key].append(func(opt_type, curr_price, strike, expiry, rf, iv, div))
                        else:
                            key=greek
                            opt_type='c'
                            results[key].append(func(opt_type, curr_price, strike, expiry, rf, iv, div))
        return pd.DataFrame(results)
    
    def implied_prob_calculation(self) -> pd.Series:
        
        """
        Calculates the implied probability of reaching each strike price by expiration. This method uses 
        the trapezoidal rule for numerical integration to compute the probabilities.
        
        Returns:
        - pd.Series: A series containing the implied probabilities for each strike.
        """
        
        call_prices = self.prices()['Call_P'].tolist()
        risk_free_rates = self.rf_rate
        days_to_expiries = self.days_to_expiry_yearformat
        diff_matrix = np.abs(self.moneyness.values[:, None] - self.moneyness.values)
        d = round(np.min(diff_matrix[diff_matrix > 0]),3)
        implied_prob = [0]  # Initialize with the first value as 0

        for i in range(1, len(call_prices) - 1):
            temp = m.exp(risk_free_rates[i] * days_to_expiries[i]) * (call_prices[i-1] + call_prices[i+1] - 2 * call_prices[i]) / (d ** 2)
            implied_prob.append(max(temp, 0))  # Use max to ensure non-negative values

        implied_prob.append(0)  # Append the last value as 0

        total_prob = sum(implied_prob)
        final_implied_prob = [prob / total_prob for prob in implied_prob]

        return pd.Series(final_implied_prob)
    
    def _main_(self)  -> pd.DataFrame:

        """
        The main method that consolidates all the calculations and returns a comprehensive dataframe. 
        This dataframe contains implied volatilities, option prices, Greeks, and implied probabilities 
        for each strike.
        
        Returns:
        - pd.DataFrame: A dataframe containing all the calculated values for each option.
        """

        # Calculate various financial metrics and concatenate them into a single DataFrame
        cal=pd.concat([self.curve_fitting_implied_vol(0.01),self.implied_prob_calculation(),self.prices(),self.Greeks()],axis=1)
        cal_col=['implied_vol_calculated','implied_prob','call_p','put_p','call_delta','put_delta',
                 'gamma','vega','theta_c','theta_p','rho_c','rho_p']
        cal.columns=cal_col
        
        # Create a DataFrame with the raw data
        data=pd.DataFrame([self.Company_name,self.Q_name,self.Current_Date,self.Expiry_Number,self.Expiry_Date,
                          self.days_to_expiry_yearformat,self.days_to_expiry,self.current_price,self.strike,self.moneyness,
                          self.implied_vol_percent,self.FN]).T
        data_col=['Company_name','Q_name','Current_Date','Expiry_Number','Expiry_Date','days_to_expiry_yearformat','days_to_expiry',
                  'current_price','strike','moneyness','imp_vol_hardcoded','FN']
        #Convert date columns to datetime format for consistency and ease of manipulation
        data.columns=data_col
        data[['Current_Date', 'Expiry_Date']] = data[['Current_Date', 'Expiry_Date']].apply(pd.to_datetime)

        # Combine the calculated metrics and raw data into a single DataFrame        
        main=pd.concat([data,cal], axis=1)
        return main
    
    def __del__(self):
        return 0


# In[4]:


class sheet_collecter(InputOutput):
    """
    The `sheet_collecter` class is designed to collect and process data from multiple sheets of an Excel file.
    It inherits from the `InputOutput` class and utilizes the `Calculate` class for processing the data.
    """
    def __init__(self,_input,company_name, folder_name, file_name,sheet_names):
        #Inheritance from InputOutput class
        super().__init__(_input,company_name, folder_name, file_name,sheet_names)
        
    def combine_three_sheets(self):
        """
        Combines and processes data from the specified sheets in the Excel file.
        
        Returns:
        - list: A list of DataFrames, each containing processed data from the respective sheet.
        """
        
        list_of_sheets=list()
        
        # Loop through each sheet name provided
        for sheet in self.sheet_names:
            # Open the Excel sheet and read its content
            read_file = self.excel_open(sheet)
            
            # Process the data using the Calculate class
            obj = Calculate(read_file)
            list_of_sheets.append(obj._main_())
            
            # Delete the Calculate object to free up memory
            del obj
        return list_of_sheets
    def extract_latest_expiry(self)-> pd.DataFrame:
        return self.combine_three_sheets()[0]
    
    def write_to_excel(self):
        
        """
        Writes data to an Excel file.
        
        Parameters:
        - sheet_collecter (list): List of DataFrames to be written to the Excel file.
        """
        
        data_output=file=f"{self._input_}/RESULT{self.company_name}/{self.folder_name}/{self.file_name}.xlsx"
        directory = os.path.dirname(data_output)
        if not os.path.exists(directory):
            os.makedirs(directory)
        with pd.ExcelWriter(data_output) as writer:
            for df, sheet in zip(self.combine_three_sheets(), self.sheet_names):
                df.to_excel(writer, sheet_name=sheet, index=False)
        
        
    def __del__(self):
        """
        Destructor for the `sheet_collecter` class.
        """
        return 0


# In[5]:


class proj_plots:
    
    """
    The proj_plots class is designed to generate and save various plots related to implied volatility 
    and implied probability. It provides functionalities to plot implied volatility and implied probability 
    curves for different expiry dates.
    
    """
    
    def __init__(self,sheets):
        
        """
        Initializes the proj_plots class with data from the provided sheets.
        
        Parameters:
        - sheets (list): A list of DataFrames containing financial data.
        """
        
        self.company_name_plot=sheets[0]['Company_name'].iloc[0]
        self.q_name_plot=sheets[0]['Q_name'].iloc[0]
        self.moneyness_plot=[sheet['moneyness'] for sheet in sheets]
        self.implied_prob_plot = [sheet['implied_prob'] for sheet in sheets]
        self.implied_vol_plot=[sheet['implied_vol_calculated'] for sheet in sheets]
        self.implied_vol_hardcoded = [sheet[sheet['imp_vol_hardcoded'] != 0]['imp_vol_hardcoded'].tolist() for sheet in sheets]
        self.moneyness_hardcoded = [sheet.loc[sheet['imp_vol_hardcoded'] != 0, 'moneyness'].tolist() for sheet in sheets]
        self.Exp_days_list = [m.ceil(sheet['days_to_expiry_yearformat'][0]*252) for sheet in sheets]
        self.Exp_list=[sheet['Expiry_Date'][0].date() for sheet in sheets]
        self.current_date=sheets[0]['Current_Date'][0].date()
        self.current_price=sheets[0]['current_price'][0]
        
    def random_colour(self):
        """
        Generates a random color.
        
        Returns:
        - tuple: A tuple representing RGB values.
        """
            
        return (random.random(), random.random(), random.random())
        
    def gen_plot_one_expiry_implied_vol(self,company_name,moneyness_plot,implied_vol_plot,implied_vol_hardcoded,moneyness_hardcoded,
                                        expiry,Exp_days_list,current_date,current_price):
        """
        Can plot the implied vol for any expiry you want with the necessary inputs. Added feature 
        for better functionality.
        
          Returns:
            fig (matplotlib.figure.Figure): The generated figure object.
        """
        plt.clf()
        fig, ax = plt.subplots()
        ax.plot(moneyness_plot, implied_vol_plot,'-o',c="black",label = f'{expiry}({Exp_days_list} day(s) left)')
        ax.plot(moneyness_hardcoded,implied_vol_hardcoded,'o',c="red",label = f'Market Values')
        ax.set_title(f"Implied volatality curve of {company_name} as on {current_date}\n Current Price = {current_price} ")
        ax.set_xlabel("Moneyness")
        ax.set_ylabel("Calculated Implied volatility")
        ax.legend()
        return fig
        
    
    def gen_plot_one_expiry_implied_prob(self,company_name,moneyness_plot,implied_prob_plot,expiry,Exp_days_list,
                                          current_date,current_price):
        """
        Can plot the implied prob for any expiry you want with the necessary inputs. Added feature 
        for better functionality.
        
          Returns:
            fig (matplotlib.figure.Figure): The generated figure object.
        """
        plt.clf()
        fig, ax = plt.subplots()
        ax.plot(moneyness_plot, implied_prob_plot, '-o',c="black",label = f'{expiry}({Exp_days_list} day(s) left)')
        ax.set_title(f"Implied probability curve of {company_name} as on {current_date}\n Current Price = {current_price} ")
        ax.set_xlabel("Moneyness")
        ax.set_ylabel("Calculated Implied probability")
        ax.legend()
        return fig
    
    def plot_all_expiries_implied_vol(self):
        """
        Plot the implied volatility curve for all expiries.

        This method generates a plot showing implied volatility against moneyness for various expiry dates 
        (depending on the number of sheets).

        Returns:
            fig (matplotlib.figure.Figure): The generated figure object.
        """
        plt.clf()
        fig,ax = plt.subplots()
        for i in range(0,len(self.moneyness_plot)):
            if(i==0):
                ax.plot(self.moneyness_plot[i], self.implied_vol_plot[i],'-o',c="black",label = f'{self.Exp_list[i]}({self.Exp_days_list[i]} day(s) left)')
                ax.plot(self.moneyness_hardcoded[i],self.implied_vol_hardcoded[i],'o',c="red",label = f'Market Values')
            else:
                ax.plot(self.moneyness_plot[i], self.implied_vol_plot[i],'-.',c=self.random_colour(),label = f'{self.Exp_list[i]}({self.Exp_days_list[i]} day(s) left)')
        ax.set_title(f"Implied volatality curve of {self.company_name_plot} as on {self.current_date}\n Current Price = {self.current_price}")
        ax.set_xlabel("Moneyness")
        ax.set_ylabel("Calculated Implied volatility")
        ax.legend()
        return fig
    
    def plot_all_expiries_implied_prob(self):
        """
        Plot the implied probability curve for all expiries.
        
        This method generates a plot showing implied probability against moneyness for various expiry dates 
        (depending on the number of sheets).
        
        Returns:
            fig (matplotlib.figure.Figure): The generated figure object.
        """
            
        plt.clf()
        fig, ax = plt.subplots()
        for i in range(0,len(self.moneyness_plot)):
            if(i==0):
                ax.plot(self.moneyness_plot[i], self.implied_prob_plot[i],'-o',c="black",label = f'{self.Exp_list[i]}({self.Exp_days_list[i]} day(s) left)')
            else:
                ax.plot(self.moneyness_plot[i], self.implied_prob_plot[i],'--',c=self.random_colour(),label = f'{self.Exp_list[i]}({self.Exp_days_list[i]} day(s) left)')
                
        ax.set_title(f"Implied probability curve of {self.company_name_plot} as on {self.current_date}\n Current Price = {self.current_price}")
        ax.set_xlabel("Moneyness")
        ax.set_ylabel("Calculated Implied probability")
        ax.legend()
        return fig
    
    def plot_latest_exp_imp_vol(self):
        #Specific function for this project
        return self.gen_plot_one_expiry_implied_vol(self.company_name_plot,self.moneyness_plot[0],self.implied_vol_plot[0],self.implied_vol_hardcoded[0],self.moneyness_hardcoded[0],
                                        self.Exp_list[0],self.Exp_days_list[0],self.current_date,self.current_price)
    def plot_latest_exp_imp_prob(self):
        #Specific function for this project
        return self.gen_plot_one_expiry_implied_prob(self.company_name_plot,self.moneyness_plot[0],self.implied_prob_plot[0],self.Exp_list[0],self.Exp_days_list[0],self.current_date,self.current_price)
    
    def save_plot(self, ax, path):
        """
        This function is an added feature to save any new fig generated for further research
        """
        ax.figure.savefig(path)
    def save_all_plots(self,path:str):
        
        """
        Save all generated plots to the specified path. Specifc function for this project

        This method saves various plots (implied volatility and implied probability) 
        for both the latest expiry and all expiries to the given directory. 
        The plots are saved in PDF format.

        Parameters:
            path (str): The base directory where the plots will be saved.

        Returns: None
        """
            
        temp_data_output=file=path+f"/PIC{self.company_name_plot}/{self.q_name_plot}/{self.current_date}/"
        temp_functions=[(self.plot_latest_exp_imp_vol,('only_latest_imp_vol')),
                        (self.plot_latest_exp_imp_prob,('only_latest_imp_prob')),
                       (self.plot_all_expiries_implied_vol,('all_exp_imp_vol')),
                        (self.plot_all_expiries_implied_prob,('all_exp_imp_prob'))]
        
        for func,name in temp_functions:
            data_output=temp_data_output+f"{name}.pdf"
            directory = os.path.dirname(data_output)
            if not os.path.exists(directory): # Creates a new directory to streamline the file-saving process
                os.makedirs(directory)
            self.save_plot(func(),data_output)
            
    def __del__(self):
        return 0


# In[6]:


#Here we are assuming that the code is being used for this project
class Strategy:
    def __init__(self,file_directory:dict):
        """
        Since we are testing out code for our specfic project, the snippet won't be dynamic for genreal use cases
        we have labeled the sheets ['BEAD','EAD','AEAD']
        in the dictioanary, incase the order of inputting sheets is changed.
        
        """
        
        """
        Initialize the Strategy object with data from the provided file directory.
        
        Parameters:
        - file_directory (dict): A dictionary containing the data sheets.
        """
        
        self.file_naming = []
        self.b_ead = []
        self._ead = []
        self.a_ead = []
        for file_key, inner_dict in file_directory.items():
            for e_key, value in inner_dict.items():
                if e_key == "BEAD":
                    self.file_naming.append(f"{value.loc[0,'Company_name']}({value.loc[0,'Current_Date'].date()})")
                    self.b_ead.append(value)
                elif e_key == 'EAD':
                    self._ead.append(value)
                else:
                    self.a_ead.append(value)
        # we are assuming here that the hard coded data has these values, if not, instead of imp_vol_hardcoded
        # input imp_vol_calculated
        self.atm_index=[b_ead.index[b_ead['moneyness'] == 1] for b_ead in self.b_ead]
        
        self.concave = [float(b_ead.loc[atm_index-1, 'imp_vol_hardcoded']) + 
                float(b_ead.loc[atm_index+1, 'imp_vol_hardcoded']) -
                2*float(b_ead.loc[atm_index, 'imp_vol_hardcoded']) 
                for b_ead, atm_index in zip(self.b_ead, self.atm_index)]

        self.current_price_b_ead = [b_ead.loc[0,'current_price'] for b_ead in self.b_ead]
        self.current_price_ead =   [_ead.loc[0,'current_price'] for _ead in self._ead]
        
        # we are assuuming for every sheet to have number of strikes, for this algo to work.Hence, the index for 
        # ATM strike is same for every sheet
        
    
    def cal_ret(self,name,current_price_b_ead,current_price_ead,concave,call_price_buy,put_price_buy,delta_put,delta_call,call_price_sell,put_price_sell):
        """
        Reference:
        Alexiou, L., A. Goyal, A. Kostakis, and L. Rompolis (2021). Pricing event risk: 
        Evidence from concave implied volatility curves. Swiss Finance Institute Research Paper (21-48).
        
        """
        """
        Calculate the straddle strategy based on the provided parameters.
        
        Returns:
        - list: A list containing the results of the straddle strategy.
        """
        
        imp_move =  (call_price_buy+put_price_buy)/ current_price_b_ead
        
        w=-(((delta_put)/put_price_buy)/((delta_call/call_price_buy)-(delta_put/put_price_buy)))
        
        straddle_return =w*(np.log(call_price_sell/call_price_buy))-((1-w)*np.log(put_price_sell/put_price_buy))
        
        price_return = np.log(current_price_ead/current_price_b_ead)
        
        result_frame= [name,price_return ,concave,imp_move , w ,straddle_return , call_price_buy, put_price_buy , 
                       delta_put ,  delta_call, call_price_sell , put_price_sell]

        return result_frame

    def list_both_strategy(self) -> list:
        """
        List the results of both the straddle and strangle strategies.
        
        Returns:
        - list: A list containing two DataFrames, one for each strategy.
        """
        temp_straddle=list()
        temp_strangle=list()
        for current_price_b_ead,current_price_ead,concave,name,atm_index,b_ead,_ead in zip(self.current_price_b_ead,
                                                                                           self.current_price_ead,
                                                                                           self.concave,
                                                                                           self.file_naming,
                                                                                           self.atm_index,
                                                                                           self.b_ead,self._ead):
            re = self.cal_ret(name,current_price_b_ead,current_price_ead,concave,
                                               float(b_ead.loc[atm_index, "call_p"]),
                                               float(b_ead.loc[atm_index, "put_p"]),
                                               float(b_ead.loc[atm_index, "put_delta"]),
                                               float(b_ead.loc[atm_index, "call_delta"]),
                                               float(_ead.loc[atm_index, "call_p"]),
                                               float(_ead.loc[atm_index, "put_p"]))
            temp_straddle.append(re)
            r1 = self.cal_ret(name,current_price_b_ead,current_price_ead,concave,
                                               float(b_ead.loc[atm_index+2, "call_p"]),
                                               float(b_ead.loc[atm_index-2, "put_p"]),
                                               float(b_ead.loc[atm_index-2, "put_delta"]),
                                               float(b_ead.loc[atm_index+2, "call_delta"]),
                                               float(_ead.loc[atm_index+2, "call_p"]),
                                               float(_ead.loc[atm_index-2, "put_p"]))
            temp_strangle.append(r1)

        straddle_df = pd.DataFrame(temp_straddle)
        c_name_straddle_df = ['name','price_return','concave','imp_move' ,'w' ,'straddle_return' , 'call_price_buy', 'put_price_buy' , 
                       'delta_put' ,  'delta_call', 'call_price_sell' , 'put_price_sell']
        straddle_df.columns=c_name_straddle_df
        straddle_df = straddle_df.set_index('name')
        strangle_df = pd.DataFrame(temp_strangle)
        c_name_strangle_df = ['name','price_return','concave','imp_move' ,'w' ,'strangle_return' , 'call_price_buy', 'put_price_buy' , 
                       'delta_put' ,  'delta_call', 'call_price_sell' , 'put_price_sell']
        strangle_df.columns=c_name_strangle_df      
        strangle_df=strangle_df.set_index('name')
        
        return [straddle_df,strangle_df]

    def save_strategy_to_excel(self, strategy_name, path):
        # Get the strategy dataframes
        temp_stad, temp_strangle = self.list_both_strategy()

        # Choose the appropriate dataframe based on the strategy_name
        strategy_df = temp_stad if strategy_name == "straddle" else temp_strangle

        directory = f"{path}/Strategy"
        data_output = os.path.join(directory, f"{strategy_name}.xlsx")

        # Create the directory if it doesn't exist
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Write the DataFrame to Excel
        with pd.ExcelWriter(data_output) as writer:
            strategy_df.to_excel(writer, index=True)


# In[8]:


"""
Used in the main function to form Dictionary data type for shortest expiry to use in Strategy function

"""
def update_nested_dict(file_name,FN, value_list, nested_dict):
    nested_dict.setdefault(file_name, {}).setdefault(FN,value_list)
    return nested_dict


# In[19]:


def main(path:str,folder_1_names,folder_2_names,individual_file_names,sheet_names):
    """
    This function:
    - Executes the entire algorithm.
    - Generates and saves the necessary plots.
    - Records the strategy results for each quarter for every company.
    - Saves the result worksheets for potential data reference.
    Note:
    - The strategy is applied only for the shortest expiry.
    - All sheet names in the raw data file must be consistent.
    
    """
    obj_dict_strat=dict()
    
    for folder_1_name,folder_2_name,individual_file_name in zip(folder_1_names,folder_2_names,individual_file_names):
        print(folder_1_name)
        obj_collecting_sheet_per_file = sheet_collecter(path,folder_1_name,folder_2_name,individual_file_name,
                                                               sheet_names)
        collecting_sheet_per_file = obj_collecting_sheet_per_file.combine_three_sheets()
        shortest_expiry = obj_collecting_sheet_per_file.extract_latest_expiry() # the latest expiry
        obj_collecting_sheet_per_file.write_to_excel() # This simply writing to files
        obj_ploting = proj_plots(collecting_sheet_per_file)
        obj_ploting.save_all_plots(path)
        plt.close('all')
        update_nested_dict(individual_file_name,shortest_expiry['FN'].iloc[0],shortest_expiry,obj_dict_strat)
    
    obj_strat = Strategy(obj_dict_strat)
    obj_strat.save_strategy_to_excel('straddle',path)
    obj_strat.save_strategy_to_excel('strangle',path)
    
    return 0


# In[17]:


# data specfic to my project 
# make sure to change naming, if it doesnot work out
path1='/Users/darshkachhara/Desktop/Project_fe800/Data' # customise this as per your address
f_1 = ['BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK','BLK',
       'BLK','BLK','BLK','BLK','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL',
       'GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','GOOGL','JPM','JPM','JPM',
       'JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM','JPM',
       'META','META','META','META','META','META','META','META','META','META','META','META','META','META','META',
       'META','META','META','META','META','META','PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP',
       'PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP','PEP','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX',
       'NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','NFLX','WBA','WBA'
       ,'WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA','WBA',
       'WBA']
file_name_for_proje = ['2021Q1BLKAPRIL14.xlsx', '2021Q1BLKAPRIL15.xlsx', '2021Q1BLKAPRIL16.xlsx', 
                       '2021Q2BLKJULY13.xlsx', '2021Q2BLKJULY14.xlsx', '2021Q2BLKJULY15.xlsx', 
                       '2021Q3BLKOCT12.xlsx', '2021Q3BLKOCT13.xlsx', '2021Q3BLKOCT14.xlsx', 
                       '2021Q4BLKJAN13.xlsx', '2021Q4BLKJAN14.xlsx', '2021Q4BLKJAN18.xlsx', 
                       '2022Q1BLKAPRIL12.xlsx', '2022Q1BLKAPRIL13.xlsx', '2022Q1BLKAPRIL14.xlsx', 
                       '2022Q2BLKJULY14.xlsx', '2022Q2BLKJULY15.xlsx', '2022Q2BLKJULY18.xlsx', 
                       '2022Q3BLKOCT12.xlsx', '2022Q3BLKOCT13.xlsx', '2022Q3BLKOCT14.xlsx', 
                       '2021Q1GOGL26APR.xlsx', '2021Q1GOGL27APR.xlsx', '2021Q1GOGL28APR.xlsx', 
                       '2021Q2GOGL26JULY.xlsx', '2021Q2GOGL27JUL.xlsx', '2021Q2GOGL28JULY.xlsx', 
                       '2021Q3GOGL25OCT.xlsx', '2021Q3GOGL26OCT.xlsx', '2021Q3GOGL27OCT.xlsx', 
                       '2021Q4GOGL31JAN.xlsx', '2021Q4GOGL1FEB.xlsx', '2021Q4GOGL2FEB.xlsx', 
                       '2022Q1GOGL25APR.xlsx', '2022Q1GOGL26APR.xlsx', '2022Q1GOGL27APR.xlsx', 
                       '2022Q2GOGL25JULY.xlsx', '2022Q2GOGL26JULY.xlsx', '2022Q2GOGL27JULY.xlsx', 
                       '2022Q3GOGL24OCT.xlsx', '2022Q3GOGL25OCT.xlsx', '2022Q3GOGL26OCT.xlsx', 
                       '2021Q1JPM13APR.xlsx', '2021Q1JPM14APR.xlsx', '2021Q1JPM15APR.xlsx', 
                       '2021Q2JPM12JULY.xlsx', '2021Q2JPM13JULY.xlsx', '2021Q2JPM14JULY.xlsx', 
                       '2021Q3JPM12OCT.xlsx', '2021Q3JPM13OCT.xlsx', '2021Q3JPM14OCT.xlsx', 
                       '2021Q4JPM13JAN.xlsx', '2021Q4JPM14JAN.xlsx', '2021Q4JPM18JAN.xlsx', 
                       '2022Q1JPM12APR.xlsx', '2022Q1JPM13APR.xlsx', '2022Q1JPM14APR.xlsx', 
                       '2022Q2JPM13JULY.xlsx', '2022Q2JPM14JULY.xlsx', '2022Q2JPM15JULY.xlsx', 
                       '2022Q3JPM13OCT.xlsx', '2022Q3JPM14OCT.xlsx', '2022Q3JPM17OCT.xlsx', 
                       '2021Q1META27APR.xlsx', '2021Q1META28APR.xlsx', '2021Q1META29APR.xlsx', 
                       '2021Q2META27JULY.xlsx', '2021Q2META28JULY.xlsx', '2021Q2META29JULY.xlsx', 
                       '2021Q3META22OCT.xlsx', '2021Q3META25OCT.xlsx', '2021Q3META26OCT.xlsx', 
                       '2021Q4META1FEB.xlsx', '2021Q4META2FEB.xlsx', '2021Q4META3FEB.xlsx', 
                       '2022Q1META26APR.xlsx', '2022Q1META27APR.xlsx', '2022Q1META28APR.xlsx', 
                       '2022Q2META26JULY.xlsx', '2022Q2META27JULY.xlsx', '2022Q2META28JULY.xlsx', 
                       '2022Q3META25OCT.xlsx', '2022Q3META26OCT.xlsx', '2022Q3META27OCT.xlsx', 
                       '2021Q1PEP14APR.xlsx', '2021Q1PEP15APR.xlsx', '2021Q1PEP16APR.xlsx', 
                       '2021Q2PEP12JULY.xlsx', '2021Q2PEP13JULY.xlsx', '2021Q2PEP14JULY.xlsx', 
                       '2021Q3PEP4OCT.xlsx', '2021Q3PEP5OCT.xlsx', '2021Q3PEP6OCT.xlsx', 
                       '2021Q4PEP9FEB.xlsx', '2021Q4PEP10FEB.xlsx', '2021Q4PEP11FEB.xlsx', 
                       '2022Q1PEP25APR.xlsx', '2022Q1PEP26APR.xlsx', '2022Q1PEP27APR.xlsx', 
                       '2022Q2PEP11JULY.xlsx', '2022Q2PEP12JULY.xlsx', '2022Q2PEP13JULY.xlsx',
                       '2022Q3PEP11OCT.xlsx', '2022Q3PEP12OCT.xlsx', '2022Q3PEP13OCT.xlsx', 
                       '2021Q1NFLX19APR.xlsx', '2021Q1NFLX20APR.xlsx', '2021Q1NFLX21APR.xlsx', 
                       '2021Q2NFLX19JULY.xlsx', '2021Q2NFLX20JULY.xlsx', '2021Q2NFLX21JULY.xlsx', 
                       '2021Q3NFLX18OCT.xlsx', '2021Q3NFLX19OCT.xlsx', '2021Q3NFLX20OCT.xlsx', 
                       '2021Q4NFLX19JAN.xlsx', '2021Q4NFLX20JAN.xlsx', '2021Q4NFLX21JAN.xlsx', 
                       '2022Q1NFLX18APR.xlsx', '2022Q1NFLX19APR.xlsx', '2022Q1NFLX20APR.xlsx', 
                       '2022Q2NFLX18JULY.xlsx', '2022Q2NFLX19JULY.xlsx', '2022Q2NFLX20JULY.xlsx', 
                       '2022Q3NFLX17OCT.xlsx', '2022Q3NFLX18OCT.xlsx', '2022Q3NFLX190CT.xlsx', 
                       '2021Q1WBA30MAR.xlsx', '2021Q1WBA31MAR.xlsx', '2021Q1WBA1APR.xlsx', '2021Q2WBA30JUNE.xlsx',
                       '2021Q2WBA1JULY.xlsx', '2021Q2WBA2JULY.xlsx', '2021Q3WBA13OCT.xlsx', '2021Q3WBA14OCT.xlsx',
                       '2021Q3WBA15OCT.xlsx', '2021Q4WBA5JAN.xlsx', '2021Q4WBA6JAN.xlsx', '2021Q4WBA7JAN.xlsx', 
                       '2022Q1WBA30MAR.xlsx', '2022Q1WBA31MAR.xlsx', '2022Q1WBA1APR.xlsx', '2022Q2WBA29JUNE.xlsx',
                       '2022Q2WBA30JUNE.xlsx', '2022Q2WBA1JULY.xlsx', '2022Q3WBA12OCT.xlsx', '2022Q3WBA13OCT.xlsx',
                       '2022Q3WBA14OCT.xlsx']
qtrs = ['2021Q1', '2021Q1', '2021Q1', '2021Q2', '2021Q2', '2021Q2', '2021Q3', '2021Q3', '2021Q3', 
        '2021Q4', '2021Q4', '2021Q4', '2022Q1', '2022Q1', '2022Q1', '2022Q2', '2022Q2', '2022Q2', 
        '2022Q3', '2022Q3', '2022Q3', '2021Q1', '2021Q1', '2021Q1', '2021Q2', '2021Q2', '2021Q2', 
        '2021Q3', '2021Q3', '2021Q3', '2021Q4', '2021Q4', '2021Q4', '2022Q1', '2022Q1', '2022Q1', '2022Q2', 
        '2022Q2', '2022Q2', '2022Q3', '2022Q3', '2022Q3', '2021Q1', '2021Q1', '2021Q1', '2021Q2', '2021Q2', 
        '2021Q2', '2021Q3', '2021Q3', '2021Q3', '2021Q4', '2021Q4', '2021Q4', '2022Q1', '2022Q1', '2022Q1', 
        '2022Q2', '2022Q2', '2022Q2', '2022Q3', '2022Q3', '2022Q3', '2021Q1', '2021Q1', '2021Q1', '2021Q2', 
        '2021Q2', '2021Q2', '2021Q3', '2021Q3', '2021Q3', '2021Q4', '2021Q4', '2021Q4', '2022Q1', '2022Q1', 
        '2022Q1', '2022Q2', '2022Q2', '2022Q2', '2022Q3', '2022Q3', '2022Q3', '2021Q1', '2021Q1', '2021Q1', 
        '2021Q2', '2021Q2', '2021Q2', '2021Q3', '2021Q3', '2021Q3', '2021Q4', '2021Q4', '2021Q4', '2022Q1', 
        '2022Q1', '2022Q1', '2022Q2', '2022Q2', '2022Q2', '2022Q3', '2022Q3', '2022Q3', '2021Q1', '2021Q1', 
        '2021Q1', '2021Q2', '2021Q2', '2021Q2', '2021Q3', '2021Q3', '2021Q3', '2021Q4', '2021Q4', '2021Q4',
        '2022Q1', '2022Q1', '2022Q1', '2022Q2', '2022Q2', '2022Q2', '2022Q3', '2022Q3', '2022Q3', '2021Q2', 
        '2021Q2', '2021Q2', '2021Q3', '2021Q3', '2021Q3', '2021Q4', '2021Q4', '2021Q4', '2022Q1', '2022Q1', 
        '2022Q1', '2022Q2', '2022Q2', '2022Q2', '2022Q3', '2022Q3', '2022Q3', '2022Q4', '2022Q4', '2022Q4']


# In[21]:


main(path1,f_1,qtrs,file_name_for_proje,['Sheet1','Sheet2','Sheet3'])

