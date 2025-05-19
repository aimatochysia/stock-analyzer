# Indonesian Stock Screener (MA-Based Approach) - Deprecated

This project is an Indonesian stock screener built using a Moving Average (MA)-based approach to help investors and traders identify potential stock opportunities. The tool analyzes stock market data from the Indonesian Stock Exchange (IDX) using simple and exponential moving averages to generate buy and sell signals, facilitating smarter investment decisions.

## ⚠️ Project Status: On Hold

The current version of the stock screener is on hold and will be deprecated. The development is transitioning to a new repository as the approach is being revamped with better, more advanced methods. The old repository will no longer be maintained, and the updated version of the screener will be moved to a separate repo.

### 📈 Current Approach

The stock screener is designed around the concept of moving averages (MA). Moving averages help identify trends, smoothing out short-term price fluctuations, and enabling better decisions based on long-term stock behavior.

Key Features of the MA-Based Approach:
- **Simple Moving Average (SMA)**: Uses the average of stock prices over a specific period to help identify the general direction of the market.
- **Crossovers**: When the short-term moving average crosses the long-term moving average, a potential buy or sell signal is generated.

### 🔧 Features

- **Stock Filtering**: Filters Indonesian stocks based on certain moving average conditions (crossovers, trends, etc.).
- **Daily updated Data**: The screener fetchespast day data from IDX (Indonesian Stock Exchange) using yfinance for accurate and up-to-date stock analysis.
- **Buy/Sell Recommendations**: Based on MA crossovers and trends, the system will generate buy and sell recommendations.

### 🛠️ Technologies Used

- **Python**: The main programming language used for the screener's backend.
- **Pandas**: For data manipulation and calculation of moving averages.
- **YFinance**: For fetching stock data (or alternative sources in development).

### 🚧 What's Next?

The current project is being moved to a new repository while we focus on exploring new approaches and techniques to refine the stock screening process. These updates may include:
- Incorporating machine learning models for better prediction and classification.
- Exploring alternative technical analysis methods like RSI, MACD, Bollinger Bands, etc.
- Creating a more user-friendly interface (UI/UX) for easier interaction with the screener.

The old repository will be **deprecated**, and the new repository will host the updated stock screener. 

### 📄 Contributing

Contributions are now welcome on the new repository. If you're interested in contributing to the development of this tool, feel free to:
- Fork the new repository.
- Create a branch for your feature or bugfix.
- Submit a pull request with your changes.

### 📬 Contact

For any inquiries or suggestions, feel free to open an issue or reach out to the project maintainers via the socials listed on the owner.

---

We appreciate your interest in this project and look forward to releasing the updated version of the stock screener in the new repository soon! Stay tuned for future updates!
