# Install dependencies
npm install

# Generate SSL certs
npx office-addin-dev-certs install

# Serve files (Terminal 1)
# Make sure you're in the project root
npx http-server -p 3000 --ssl -c-1 --cert C:\Users\User\.office-addin-dev-certs\localhost.crt --key C:\Users\User\.office-addin-dev-certs\localhost.key

# Start backend (Terminal 2)
cd backend && python app.py

# Sideload add-in (Terminal 3, one-time)
npm start