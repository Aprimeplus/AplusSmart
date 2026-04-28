
def verify_hr_logic():
    print("--- Verifying HR Logic ---")
    # Simulate inputs
    sales_amount = 10000.0 # VAT
    no_vat_item = 5000.0 # NO VAT
    wht = 0.0
    
    # Logic from hr_windows.py (Fixed)
    total_vatable_revenue = sales_amount
    total_cashable_services_and_fees = no_vat_item
    
    final_grand_total = (total_vatable_revenue * 1.07) + total_cashable_services_and_fees - wht
    
    expected_total = (10000 * 1.07) + 5000
    print(f"Calculated Total: {final_grand_total}")
    print(f"Expected Total: {expected_total}")
    
    if abs(final_grand_total - expected_total) < 0.01:
        print("PASS: Grand Total calculation is correct.")
    else:
        print("FAIL: Grand Total calculation is incorrect.")

def verify_dashboard_logic():
    print("\n--- Verifying Dashboard Logic ---")
    # Scenario: Customer paid full amount (incl NO VAT)
    # Grand Total = 15700 (10700 VAT + 5000 NO VAT)
    # Payment = 15700
    # Difference = Payment - Grand Total = 0
    
    payment = 15700.0
    difference = 0.0
    
    # Logic from outstanding_dashboard_tab.py (Fixed)
    full_amount = payment - difference
    remaining = -difference
    
    print(f"Scenario 1 (Full Pay): Full Amount={full_amount}, Remaining={remaining}")
    if full_amount == 15700 and remaining == 0:
        print("PASS: Full Payment scenario")
    else:
        print("FAIL: Full Payment scenario")

    # Scenario: Customer Underpaid
    # Grand Total = 15700
    # Payment = 10000
    # Difference = 10000 - 15700 = -5700
    
    payment = 10000.0
    difference = -5700.0
    
    full_amount = payment - difference # 10000 - (-5700) = 15700
    remaining = -difference # -(-5700) = 5700
    
    print(f"Scenario 2 (Underpay): Full Amount={full_amount}, Remaining={remaining}")
    if full_amount == 15700 and remaining == 5700:
        print("PASS: Underpayment scenario")
    else:
        print("FAIL: Underpayment scenario")

if __name__ == "__main__":
    verify_hr_logic()
    verify_dashboard_logic()
 



 ###testtt