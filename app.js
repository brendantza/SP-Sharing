document.addEventListener('DOMContentLoaded', () => {
    const apiKeyInput = document.getElementById('api-key');
    const submitBtn = document.getElementById('submit-key');
    const filterBtn = document.getElementById('filter-btn');
    const transactionList = document.getElementById('transactions');
    const ctx = document.getElementById('transaction-chart').getContext('2d');
    let chart;

    let transactions = [];
    let currentApiKey = '';

    submitBtn.addEventListener('click', () => {
        currentApiKey = apiKeyInput.value;
        if (currentApiKey) {
            const thirtyDaysAgo = new Date();
            thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
            document.getElementById('start-date').valueAsDate = thirtyDaysAgo;
            document.getElementById('end-date').valueAsDate = new Date();
            fetchTransactions(currentApiKey);
        } else {
            alert('Please enter your API key.');
        }
    });

    filterBtn.addEventListener('click', () => {
        if (currentApiKey) {
            renderAll(currentApiKey);
        } else {
            alert('Please enter your API key first.');
        }
    });

    // Add keyboard listeners for quicker actions
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && document.activeElement === apiKeyInput) {
            e.preventDefault();
            if (currentApiKey) {
                renderAll(currentApiKey);
            } else {
                submitBtn.click();
            }
        }
    });

    document.getElementById('time-view').addEventListener('change', () => {
        if (currentApiKey) {
            renderAll(currentApiKey);
        }
    });

    function renderAll(apiKey) {
        const filtered = filterData(transactions);
        renderChart(filtered);
        renderTransactions(filtered);
    }

    async function fetchTransactions(apiKey) {
        const startDateInput = document.getElementById('start-date').value;
        const endDateInput = document.getElementById('end-date').value;

        let fromTimestamp = '';
        let toTimestamp = '';

        if (startDateInput && endDateInput) {
            // Ensure end date includes the full day
            const endDate = new Date(endDateInput);
            endDate.setHours(23, 59, 59);
            fromTimestamp = Math.floor(new Date(startDateInput).getTime() / 1000);
            toTimestamp = Math.floor(endDate.getTime() / 1000);
        }

        // Try fetching without any date constraints first
        const url = `https://api.torn.com/user/?selections=log&key=${apiKey}`;

        console.log('Fetching from:', new Date(fromTimestamp * 1000).toLocaleDateString(), 'to:', new Date(toTimestamp * 1000).toLocaleDateString());
        console.log('API URL:', url);

        try {
            const response = await fetch(url);
            const data = await response.json();

            console.log('API Response Keys:', Object.keys(data.log || {}).length);
            console.log('First few raw log entries:', Object.values(data.log || {}).slice(0, 5).map(entry => ({
                log: entry.log,
                timestamp: new Date(entry.timestamp * 1000).toLocaleDateString(),
                title: entry.title,
                data: entry.data
            })));

            // Get all timestamps to see the actual range
            if (data.log) {
                const timestamps = Object.values(data.log).map(entry => entry.timestamp);
                const minTimestamp = Math.min(...timestamps);
                const maxTimestamp = Math.max(...timestamps);
                console.log('Raw data timestamp range:',
                    new Date(minTimestamp * 1000).toLocaleDateString(),
                    'to',
                    new Date(maxTimestamp * 1000).toLocaleDateString()
                );
            }

            if (data.error) {
                alert(`Error: ${data.error.error}`);
                return;
            }
            if (!data.log) {
                alert('No log data found in the API response for the selected date range.');
                transactions = []; // Clear transactions if no data is found
            } else {
                transactions = processLogData(data.log);
                console.log('Processed transactions:', transactions.length);
                console.log('Date range of transactions:', transactions.map(tx => new Date(tx.timestamp * 1000).toLocaleDateString()));
            }
            renderChart(transactions);
            renderTransactions(transactions);
        } catch (error) {
            console.error('Error fetching data:', error);
            alert('Failed to fetch data from the Torn API.');
        }
    }

    function processLogData(log) {
        const processedTransactions = [];
        for (const key in log) {
            const logEntry = log[key];
            const title = logEntry.title.toLowerCase();
            let amount = 0;
            let type = '';

            // Handle money gains from crimes
            if (logEntry.log === 9015 && logEntry.data && logEntry.data.money_gained) {
                amount = logEntry.data.money_gained;
                type = 'incoming';
            }
            // Handle company pay
            else if (logEntry.log === 6221 && logEntry.data && logEntry.data.money) {
                amount = logEntry.data.money;
                type = 'incoming';
            }
            // Handle racing
            else if (logEntry.log === 8410 && logEntry.data && logEntry.data.money) {
                amount = logEntry.data.money > 0 ? logEntry.data.money : -Math.abs(logEntry.data.money);
                type = amount > 0 ? 'incoming' : 'outgoing';
            }
            // Handle casino
            else if ((logEntry.log === 8210 || logEntry.log === 8220 || logEntry.log === 8230 || logEntry.log === 8240) && logEntry.data && logEntry.data.money) {
                amount = logEntry.data.money;
                type = amount > 0 ? 'incoming' : 'outgoing';
            }
            // Handle faction payments/donations
            else if (logEntry.log === 7610 && logEntry.data && logEntry.data.money) {
                amount = -logEntry.data.money;
                type = 'outgoing';
            }
            // Handle referrals
            else if (logEntry.log === 9310 && logEntry.data && logEntry.data.money) {
                amount = logEntry.data.money;
                type = 'incoming';
            }
            // Handle market sales
            else if (logEntry.log === 1210 && logEntry.data && logEntry.data.cost_total) {
                amount = logEntry.data.cost_total;
                type = 'incoming';
            }
            // Handle bazaar sales
            else if (logEntry.log === 8511 && logEntry.data && logEntry.data.cost) {
                amount = logEntry.data.cost;
                type = 'incoming';
            }
            // Handle stock dividends
            else if (logEntry.log === 2510 && logEntry.data && logEntry.data.money) {
                amount = logEntry.data.money;
                type = 'incoming';
            }
            // Handle auction purchases
            else if (logEntry.log === 4110 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle education
            else if (logEntry.log === 7110 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle medical
            else if (logEntry.log === 7310 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle points earnings
            else if (logEntry.log === 9710 && logEntry.data && logEntry.data.money) {
                amount = Math.abs(logEntry.data.money);
                type = 'incoming';
            }
            // Handle item market purchases
            else if (logEntry.log === 1310 && logEntry.data && logEntry.data.cost_total) {
                amount = -logEntry.data.cost_total;
                type = 'outgoing';
            }
            // Handle sending money
            else if (logEntry.log === 8010 && logEntry.data && logEntry.data.money) {
                amount = -logEntry.data.money;
                type = 'outgoing';
            }
            // Handle bazaar purchases
            else if (logEntry.log === 8510 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle property charges
            else if (logEntry.log === 4101 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle reviving someone from jail
            else if (logEntry.log === 7120 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle garage fees
            else if (logEntry.log === 4510 && logEntry.data && logEntry.data.cost) {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }
            // Handle abroad items
            else if (logEntry.log === 4910 && logEntry.data && logEntry.data.cost) {
                amount = logEntry.data.cost;
                type = 'incoming';
            }
            // Handle loan repayments or interest (assuming positive for interest earned)
            else if (logEntry.log === 7510 && logEntry.data && logEntry.data.money) {
                amount = logEntry.data.money;
                type = amount > 0 ? 'incoming' : 'outgoing';
            }
            // Catch-all for other financial transactions with money field
            else if (logEntry.data && logEntry.data.money && typeof logEntry.data.money === 'number') {
                amount = logEntry.data.money;
                type = amount > 0 ? 'incoming' : 'outgoing';
            }
            else if (logEntry.data && logEntry.data.cost && typeof logEntry.data.cost === 'number') {
                amount = -logEntry.data.cost;
                type = 'outgoing';
            }


            if (type) {
                processedTransactions.push({
                    timestamp: logEntry.timestamp,
                    type: type,
                    amount: amount,
                    title: logEntry.title
                });
            }
        }
        return processedTransactions;
    }

    function renderTransactions(data) {
        transactionList.innerHTML = '';
        const filteredData = filterData(data);
        const groupedByDay = {};

        filteredData.forEach(tx => {
            const date = new Date(tx.timestamp * 1000).toLocaleDateString();
            if (!groupedByDay[date]) {
                groupedByDay[date] = {
                    transactions: [],
                    total: 0
                };
            }
            groupedByDay[date].transactions.push(tx);
            groupedByDay[date].total += tx.amount;
        });

        for (const date in groupedByDay) {
            const day = groupedByDay[date];
            const dateLi = document.createElement('li');
            dateLi.innerHTML = `<strong>${date} - Total: $${day.total.toLocaleString()}</strong>`;
            transactionList.appendChild(dateLi);

            const childUl = document.createElement('ul');
            day.transactions.forEach(tx => {
                const txLi = document.createElement('li');
                const transactionDate = new Date(tx.timestamp * 1000).toLocaleString();
                const formattedAmount = tx.amount.toLocaleString();
                const direction = tx.amount > 0 ? 'Incoming' : 'Outgoing';
                txLi.innerHTML = `<strong>${transactionDate}</strong> - <em>${tx.title}</em> (${direction}) - <strong>$${formattedAmount}</strong>`;
                childUl.appendChild(txLi);
            });
            transactionList.appendChild(childUl);
        }
    }

    function renderChart(data) {
        const filteredData = filterData(data);
        const aggregatedData = aggregateData(filteredData);

        const labels = Object.keys(aggregatedData).sort();
        const incoming = labels.map(date => aggregatedData[date].incoming);
        const outgoing = labels.map(date => Math.abs(aggregatedData[date].outgoing));

        if (chart) {
            chart.destroy();
        }

        if (labels.length === 0) {
            // No data to display
            chart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: ['No Data'],
                    datasets: [{
                        label: 'No transactions found in the selected range.',
                        data: [0],
                        backgroundColor: 'rgba(128, 128, 128, 0.2)',
                        borderColor: 'rgba(128, 128, 128, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    },
                    plugins: {
                        legend: {
                            display: false
                        }
                    }
                }
            });
            return;
        }

        chart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Incoming',
                    data: incoming,
                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                    borderColor: 'rgba(75, 192, 192, 1)',
                    borderWidth: 1
                }, {
                    label: 'Outgoing',
                    data: outgoing,
                    backgroundColor: 'rgba(255, 99, 132, 0.2)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
    }

    function filterData(data) {
        const startDate = new Date(document.getElementById('start-date').value);
        const endDate = new Date(document.getElementById('end-date').value);
        const transactionType = document.getElementById('transaction-type').value;

        return data.filter(tx => {
            const txDate = new Date(tx.timestamp * 1000);
            const isAfterStartDate = !document.getElementById('start-date').value || txDate >= startDate;
            const isBeforeEndDate = !document.getElementById('end-date').value || txDate <= endDate;
            const isCorrectType = transactionType === 'all' || (transactionType === 'incoming' && tx.amount > 0) || (transactionType === 'outgoing' && tx.amount < 0);
            return isAfterStartDate && isBeforeEndDate && isCorrectType;
        });
    }

    function aggregateData(data) {
        const timeView = document.getElementById('time-view').value;
        const totals = {};
        data.forEach(tx => {
            let key = '';
            const date = new Date(tx.timestamp * 1000);
            if (timeView === 'day') {
                key = date.toLocaleDateString();
            } else if (timeView === 'week') {
                const year = date.getFullYear();
                const week = Math.ceil((date - new Date(year, 0, 1)) / (7 * 24 * 60 * 60 * 1000));
                key = `${year}-W${week.toString().padStart(2, '0')}`;
            } else if (timeView === 'month') {
                const year = date.getFullYear();
                const month = (date.getMonth() + 1).toString().padStart(2, '0');
                key = `${year}-${month}`;
            }

            if (!totals[key]) {
                totals[key] = { incoming: 0, outgoing: 0 };
            }
            if (tx.amount > 0) {
                totals[key].incoming += tx.amount;
            } else {
                totals[key].outgoing += tx.amount;
            }
        });
        return totals;
    }
});
