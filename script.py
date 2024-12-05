# Import the necessary libraries
import requests

# Set the URL of the server from which to extract data
url = 'https://www.example.com/api/data'

# Make a GET request to the server
response = requests.get(url)

# Check the response status code
if response.status_code == 200:
  # Extract the data from the response
  data = response.json()

  # Print the data
  print(data)
else:
  print('Failed to extract data from the server.')
