#!/bin/bash

# Prompt user for principal amount
read -p "Enter the principal amount: " principal

# Prompt user for interest rate
read -p "Enter the annual interest rate (in percentage): " rate

# Prompt user for time period
read -p "Enter the time period (in years): " time

# Convert interest rate percentage to decimal
rate_decimal=$(echo "scale=4; $rate / 100" | bc)

# Calculate simple interest
simple_interest=$(echo "scale=2; $principal * $rate_decimal * $time" | bc)

# Display the result
echo "Simple Interest: $simple_interest"
