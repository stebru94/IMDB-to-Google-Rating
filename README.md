<h1>Script Purpose</h1>
This Python script is designed for movie enthusiasts looking <b>to discover highly acclaimed and audience-appreciated movies.</b> Its main goal is to provide users with a way to create their own movie database based on criteria such as rating, release year, genre, and duration, using information gathered from IMDb and Google Rating.

<h1>How It Works</h1>
The underlying idea of the script is to <b>use IMDb as the primary source to obtain movie information</b>, such as <b>title, release year, genre</b>, and <b>duration</b>. Then, it uses Google Rating to find out the overall rating of the movie and how much it was liked by users.

<h1>Analysis Goal</h1>
With the collected data, users can easily create customized lists of movies that match their tastes and preferences. For example, one can create a list of the most appreciated movies of a specific genre, country, or decade. This allows users to select and watch movies based on their specific preferences, thereby avoiding low-quality content and discovering hidden gems.

<h1>Configuration and Usage</h1>
Before running the script, you need to configure the config.ini file with the desired preferences, such as the name of the Excel file to save the data, the IMDb URL to fetch the data from, and other options. After configuration, running the script will automatically start collecting and saving movie data in the specified Excel file.

<h1>Web Driver Requirements </h1>
The script uses Selenium WebDriver to interact with the browser and gather data from IMDb and Google Rating. Make sure you have the correct driver for your browser and have it set up in the specified path in the Python code.

<h1>Example Usage</h1>
After running the script, a browser will be launched that navigates to IMDb and collects movie data. This data will then be processed to obtain the Google rating and saved in the specified Excel file. Users can then review the data and select movies to watch based on their preferences.
