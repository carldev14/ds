print('\nCalculate VAT\n')
totalAmount = float(input('Enter the total amount: '))
VAT_rate = 1.12  # Assuming VAT rate is 12%, so 1 + VAT_rate = 1.12
result = totalAmount * VAT_rate - totalAmount
rounded_result = round(result, 2)
print(f"VAT amount: {rounded_result}")
