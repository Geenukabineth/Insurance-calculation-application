from typing import Union, List, Any

import numpy as np

class CashFlow(object):
    def __init__(self, claim_amount: float, number_of_years: int, discount_rate: Union[float, List], payment_pattern: Any = None) -> None:
        self.claim_amount = claim_amount
        self.number_of_years = number_of_years
        self.discount_rate = discount_rate
        self.payment_pattern = payment_pattern
        
        self.payment_years = self.get_payment_years()
        self.payment_pattern_ = self.get_payment_pattern()
        self.estimated_payment = self.get_estimated_payments()
        self.discount_rate_array = self.get_discount_rates()        
        self.discount_factor_array = self.get_dicount_factors()
        self.discounted_estimated_payment = self.get_discounted_estimated_payment()
        
        self.net_present_value = np.sum(self.discounted_estimated_payment)
        self.interest_accretion_array = self.get_interest_accretion_array()
        self.total_interest_accretion = np.sum(self.interest_accretion_array)
        
        self.finance_income_ = self.total_interest_accretion - self.net_present_value
    
    def get_payment_years(self):
        payement_years: np.ndarray = np.arange(1, self.number_of_years+1)
        
        return payement_years
    
    def get_payment_pattern(self):
        try:
            if self.payment_pattern is None:
                payment_pattern = np.full(shape=self.number_of_years, fill_value=1/self.number_of_years)
                return payment_pattern
            else:
                if isinstance(self.payment_pattern, list):
                    payment_pattern = self.payment_pattern
                    return payment_pattern

                else:
                    raise ValueError("Payment Pattern Should be a List")
                
        except ValueError as error:
            raise error
        
    def get_estimated_payments(self):
        return self.payment_pattern_ * self.claim_amount
        
    def get_discount_rates(self):
        if isinstance(self.discount_rate, float):
            discount_rate_array = np.full(shape=self.number_of_years, fill_value=self.discount_rate)
            return discount_rate_array
        else:
            discount_rate_array = self.discount_rate
            return discount_rate_array
        
    def get_dicount_factors(self):
        discount_factors = []
        for year in range(self.number_of_years):
            period = self.payment_years[year]
            discount_rate = self.discount_rate_array[year]
            
            discount_factor = self.calculate_dicounting_factor(period=period, discount_rate=discount_rate)
            discount_factors.append(discount_factor)
            
        return np.array(discount_factors)
    
    def get_discounted_estimated_payment(self):
        return self.estimated_payment * self.discount_factor_array
    
    def get_interest_accretion_array(self):
        interest_accretion_array = []
        
        for year in range(self.number_of_years):
            discounted_cashflow = self.discounted_estimated_payment[year]
            discount_rate = self.discount_rate_array[year]
            
            interest_accretion = self.interest_accretion(discounted_cashflow=discounted_cashflow, discount_rate=discount_rate)
            interest_accretion_array.append(interest_accretion)
            
        return np.array(interest_accretion_array)
    
    @staticmethod
    def calculate_dicounting_factor(period: int, discount_rate: float):
        return 1/(1+discount_rate)**period
    
    @staticmethod
    def interest_accretion(discounted_cashflow: float, discount_rate: float):
        return discounted_cashflow * (1 + discount_rate)
    
    
def main():
    claim_amount = 10000
    discount_rate = 0.1
    number_of_years = 4
    
    model = CashFlow(claim_amount=claim_amount, number_of_years=number_of_years, discount_rate=discount_rate)
    
    print(model.payment_years)
    print(model.payment_pattern_)
    print(model.estimated_payment)
    print(model.discount_factor_array)
    print(model.discounted_estimated_payment)
    print("----------------------------------------")
    print(model.net_present_value)
    print("----------------------------------------")
    print(model.payment_years)
    print(model.discounted_estimated_payment)
    print(model.interest_accretion_array)
    print("----------------------------------------")
    print(model.total_interest_accretion)
    print("----------------------------------------")
    print(model.finance_income_)
    
    

if __name__ == "__main__":
    main() 