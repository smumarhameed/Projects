import streamlit as st

def celsius_to_fahrenheit(celsius_temp):
    return (celsius_temp * 9/5) + 32

def fahrenheit_to_celsius(fahrenheit_temp):
    return (fahrenheit_temp - 32) * 5/9

def main():
    st.title("Temperature Converter")
    st.sidebar.title("Options")

    conversion_type = st.sidebar.selectbox(
        "Select conversion type",
        ("Celsius to Fahrenheit", "Fahrenheit to Celsius")
    )

    if conversion_type == "Celsius to Fahrenheit":
        celsius_temp = st.number_input("Enter temperature in Celsius", value=0.0)
        converted_temp = celsius_to_fahrenheit(celsius_temp)
        st.write(f"{celsius_temp} Celsius is equal to {converted_temp} Fahrenheit.")

    elif conversion_type == "Fahrenheit to Celsius":
        fahrenheit_temp = st.number_input("Enter temperature in Fahrenheit", value=32.0)
        converted_temp = fahrenheit_to_celsius(fahrenheit_temp)
        st.write(f"{fahrenheit_temp} Fahrenheit is equal to {converted_temp} Celsius.")

if __name__ == "__main__":
    main()
