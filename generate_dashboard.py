#!/usr/bin/env python3
"""
Dashboard Generator for Amazon Vacuum Scraper
Reads Excel data and generates interactive HTML dashboard
"""

import glob
import json
import math
import os
import sys
from datetime import datetime

import pandas as pd

import config

DASHBOARD_FILENAME = 'amazon_dashboard.html'

def find_latest_excel_file(output_dir, base_filename):
    """
    Locate the most recently modified output file. Scraper runs are saved
    with an ISO week stamp inserted before the extension (e.g.
    Amazon_Vacuum_Cleaners_Filtered_2026-W35.xlsx), so a fixed filename
    won't match -- glob for the pattern and pick the newest by mtime.
    """
    name, ext = os.path.splitext(base_filename)
    pattern = os.path.join(output_dir, f"{name}_*{ext}")
    candidates = glob.glob(pattern)

    # Fall back to the unstamped name for backwards compatibility.
    plain_path = os.path.join(output_dir, base_filename)
    if os.path.exists(plain_path):
        candidates.append(plain_path)

    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)

def read_excel_data(file_path):
    """Read Excel file and return dataframe"""
    try:
        df = pd.read_excel(file_path, sheet_name='All Products')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def calculate_metrics(df):
    """Calculate all metrics from the dataframe"""
    if df is None or df.empty:
        print("No data to calculate metrics")
        return None

    metrics = {
        "collection_date": datetime.now().strftime('%Y-%m-%d'),
        "total_products": len(df),
        "price_by_brand": {},
        "rating_by_brand": {},
        "rating_distribution": {
            "5-star": 0,
            "4-star": 0,
            "3-star": 0,
            "below-3": 0
        },
        "fulfilled_by": {
            "Amazon - FREE Shipping": 0,
            "Amazon": 0,
            "Not Available": 0
        },
        "top_rated": [],
        "most_reviewed": [],
        "best_value": []
    }

    # Price and Rating are already numeric columns in the sheet (the
    # scraper stores the raw numeric price, not a formatted "AED1,031.00"
    # string), and "N/A" cells come through pandas as NaN automatically.
    # Brand is its own column too -- no need to re-derive it from Title.
    brand_series = df['Brand'].astype(str).str.replace('‎', '', regex=False).str.strip()
    brand_series = brand_series.where(df['Brand'].notna(), 'Unknown')

    for brand, brand_data in df.groupby(brand_series):
        price_data = brand_data['Price'].dropna()
        if not price_data.empty:
            metrics['price_by_brand'][brand] = {
                "mean": round(price_data.mean(), 2),
                "count": len(brand_data)
            }

        rating_data = brand_data['Rating'].dropna()
        if not rating_data.empty:
            metrics['rating_by_brand'][brand] = {
                "mean": round(rating_data.mean(), 2),
                "count": len(rating_data)
            }

    # Calculate rating distribution (skip products with no rating at all --
    # they're missing data, not a "below 3-star" rating)
    for rating in df['Rating']:
        if rating is None or (isinstance(rating, float) and math.isnan(rating)):
            continue
        rating = float(rating)
        if rating >= 5.0:
            metrics['rating_distribution']['5-star'] += 1
        elif rating >= 4.0:
            metrics['rating_distribution']['4-star'] += 1
        elif rating >= 3.0:
            metrics['rating_distribution']['3-star'] += 1
        else:
            metrics['rating_distribution']['below-3'] += 1

    # Calculate fulfillment methods
    for idx, row in df.iterrows():
        fulfilled_by = str(row.get('Fulfilled by', '')).strip()
        if 'FREE' in fulfilled_by.upper():
            metrics['fulfilled_by']['Amazon - FREE Shipping'] += 1
        elif 'Amazon' in fulfilled_by:
            metrics['fulfilled_by']['Amazon'] += 1
        else:
            metrics['fulfilled_by']['Not Available'] += 1

    # Get top rated products (5.0 rating first, then by review count)
    df_clean = df.dropna(subset=['Rating', 'Price']).copy()

    top_rated = df_clean.nlargest(10, 'Rating')
    metrics['top_rated'] = [
        {
            "Title": row['Title'],
            "Price": row['Price'] if pd.notna(row['Price']) else None,
            "Rating": float(row['Rating']),
            "Reviews": int(row['Reviews']) if pd.notna(row['Reviews']) else 0
        }
        for idx, row in top_rated.iterrows()
    ]

    # Get most reviewed products
    most_reviewed = df_clean.nlargest(10, 'Reviews')
    metrics['most_reviewed'] = [
        {
            "Title": row['Title'],
            "Price": row['Price'] if pd.notna(row['Price']) else None,
            "Rating": float(row['Rating']),
            "Reviews": int(row['Reviews']) if pd.notna(row['Reviews']) else 0
        }
        for idx, row in most_reviewed.iterrows()
    ]

    # Get best value products (high rating + low price)
    if not df_clean.empty:
        # Normalize price and rating for scoring
        min_price = df_clean['Price'].min()
        max_price = df_clean['Price'].max()
        price_range = max_price - min_price

        if price_range > 0:
            df_clean['Price_Score'] = 1 - ((df_clean['Price'] - min_price) / price_range)
        else:
            df_clean['Price_Score'] = 1.0
        df_clean['Rating_Score'] = df_clean['Rating'] / 5.0
        df_clean['Value_Score'] = (df_clean['Price_Score'] + df_clean['Rating_Score']) / 2

        best_value = df_clean.nlargest(10, 'Value_Score')
        metrics['best_value'] = [
            {
                "Title": row['Title'],
                "Price": row['Price'],
                "Rating": float(row['Rating'])
            }
            for idx, row in best_value.iterrows()
        ]

    return metrics

def generate_html(metrics):
    """Generate HTML dashboard from metrics"""

    # Calculate summary statistics
    prices_numeric = []
    ratings_numeric = []
    for brand_data in metrics['price_by_brand'].values():
        prices_numeric.append(brand_data['mean'])
    for brand_data in metrics['rating_by_brand'].values():
        ratings_numeric.append(brand_data['mean'])

    avg_price = sum(prices_numeric) / len(prices_numeric) if prices_numeric else 0
    avg_rating = sum(ratings_numeric) / len(ratings_numeric) if ratings_numeric else 0

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Amazon Vacuum Analytics Dashboard</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}

        .container {{
            max-width: 1400px;
            margin: 0 auto;
        }}

        header {{
            background: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }}

        h1 {{
            color: #333;
            font-size: 32px;
            margin-bottom: 10px;
        }}

        .header-info {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: #666;
            font-size: 14px;
        }}

        .kpi-row {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}

        .kpi-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            border-left: 4px solid #667eea;
        }}

        .kpi-label {{
            color: #999;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
        }}

        .kpi-value {{
            font-size: 28px;
            font-weight: bold;
            color: #333;
        }}

        .kpi-subtext {{
            color: #999;
            font-size: 12px;
            margin-top: 8px;
        }}

        .charts-row {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}

        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }}

        .chart-title {{
            font-size: 16px;
            font-weight: 600;
            color: #333;
            margin-bottom: 20px;
        }}

        .chart-container {{
            position: relative;
            height: 300px;
        }}

        .table-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 30px;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }}

        thead {{
            background: #f8f9fa;
            border-bottom: 2px solid #667eea;
        }}

        th {{
            padding: 12px;
            text-align: left;
            font-weight: 600;
            color: #333;
            font-size: 13px;
            text-transform: uppercase;
        }}

        td {{
            padding: 12px;
            border-bottom: 1px solid #eee;
            color: #666;
            font-size: 14px;
        }}

        tr:hover {{
            background: #f8f9fa;
        }}

        .price-badge {{
            background: #e3f2fd;
            color: #1976d2;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }}

        .rating-badge {{
            background: #fff3e0;
            color: #f57c00;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }}

        .tabs {{
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            border-bottom: 2px solid #eee;
        }}

        .tab-button {{
            background: none;
            border: none;
            padding: 12px 20px;
            font-size: 14px;
            font-weight: 600;
            color: #999;
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: all 0.3s ease;
        }}

        .tab-button.active {{
            color: #667eea;
            border-bottom-color: #667eea;
        }}

        .tab-content {{
            display: none;
        }}

        .tab-content.active {{
            display: block;
        }}

        .grid-2col {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 20px;
        }}

        @media (max-width: 768px) {{
            .charts-row,
            .grid-2col {{
                grid-template-columns: 1fr;
            }}

            h1 {{
                font-size: 24px;
            }}

            .header-info {{
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
            }}
        }}

        .last-updated {{
            text-align: center;
            color: #999;
            font-size: 12px;
            margin-top: 40px;
            padding: 20px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 Amazon Vacuum Analytics Dashboard</h1>
            <div class="header-info">
                <span>Last Updated: {metrics['collection_date']}</span>
                <span>Data Collection Week: W{datetime.strptime(metrics['collection_date'], '%Y-%m-%d').isocalendar()[1]}</span>
            </div>
        </header>

        <!-- KPI Cards -->
        <div class="kpi-row">
            <div class="kpi-card">
                <div class="kpi-label">Total Products</div>
                <div class="kpi-value">{metrics['total_products']}</div>
                <div class="kpi-subtext">Products collected</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Average Price</div>
                <div class="kpi-value">AED {avg_price:,.0f}</div>
                <div class="kpi-subtext">Across all brands</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Average Rating</div>
                <div class="kpi-value">{avg_rating:.2f} ⭐</div>
                <div class="kpi-subtext">Customer satisfaction</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Total Reviews</div>
                <div class="kpi-value">{sum(p.get('Reviews', 0) for p in metrics['top_rated'])}</div>
                <div class="kpi-subtext">Across top products</div>
            </div>
        </div>

        <!-- Charts Section -->
        <div class="charts-row">
            <div class="chart-card">
                <div class="chart-title">💰 Price Distribution by Brand</div>
                <div class="chart-container">
                    <canvas id="priceChart"></canvas>
                </div>
            </div>

            <div class="chart-card">
                <div class="chart-title">⭐ Average Rating by Brand</div>
                <div class="chart-container">
                    <canvas id="ratingChart"></canvas>
                </div>
            </div>
        </div>

        <div class="charts-row">
            <div class="chart-card">
                <div class="chart-title">📦 Rating Distribution</div>
                <div class="chart-container">
                    <canvas id="ratingDistChart"></canvas>
                </div>
            </div>

            <div class="chart-card">
                <div class="chart-title">🚚 Fulfillment Method</div>
                <div class="chart-container">
                    <canvas id="fulfillmentChart"></canvas>
                </div>
            </div>
        </div>

        <!-- Top Products Tabs -->
        <div class="table-card">
            <h3 style="font-size: 18px; margin-bottom: 20px;">🏆 Top Products</h3>

            <div class="tabs">
                <button class="tab-button active" onclick="switchTab('top-rated')">Top Rated</button>
                <button class="tab-button" onclick="switchTab('most-reviewed')">Most Reviewed</button>
                <button class="tab-button" onclick="switchTab('best-value')">Best Value</button>
            </div>

            <!-- Top Rated Tab -->
            <div id="top-rated" class="tab-content active">
                <table>
                    <thead>
                        <tr>
                            <th>Product Title</th>
                            <th>Price (AED)</th>
                            <th>Rating</th>
                            <th>Reviews</th>
                        </tr>
                    </thead>
                    <tbody id="topRatedTable">
                    </tbody>
                </table>
            </div>

            <!-- Most Reviewed Tab -->
            <div id="most-reviewed" class="tab-content">
                <table>
                    <thead>
                        <tr>
                            <th>Product Title</th>
                            <th>Price (AED)</th>
                            <th>Rating</th>
                            <th>Reviews</th>
                        </tr>
                    </thead>
                    <tbody id="mostReviewedTable">
                    </tbody>
                </table>
            </div>

            <!-- Best Value Tab -->
            <div id="best-value" class="tab-content">
                <table>
                    <thead>
                        <tr>
                            <th>Product Title</th>
                            <th>Price (AED)</th>
                            <th>Rating</th>
                        </tr>
                    </thead>
                    <tbody id="bestValueTable">
                    </tbody>
                </table>
            </div>
        </div>

        <div class="last-updated">
            Dashboard updates every Monday at 8 AM Dubai time
        </div>
    </div>

    <script>
        const dashboardData = {json.dumps(metrics)};

        // Price Chart
        const priceCtx = document.getElementById('priceChart').getContext('2d');
        new Chart(priceCtx, {{
            type: 'bar',
            data: {{
                labels: Object.keys(dashboardData.price_by_brand),
                datasets: [{{
                    label: 'Average Price (AED)',
                    data: Object.values(dashboardData.price_by_brand).map(x => x.mean),
                    backgroundColor: ['#667eea', '#764ba2', '#f093fb', '#4facfe', '#ff6b6b'],
                    borderRadius: 5
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ y: {{ beginAtZero: true }} }}
            }}
        }});

        // Rating Chart
        const ratingCtx = document.getElementById('ratingChart').getContext('2d');
        new Chart(ratingCtx, {{
            type: 'bar',
            data: {{
                labels: Object.keys(dashboardData.rating_by_brand),
                datasets: [{{
                    label: 'Average Rating',
                    data: Object.values(dashboardData.rating_by_brand).map(x => x.mean),
                    backgroundColor: ['#ffd700', '#ffa500', '#ff69b4', '#32cd32', '#1e90ff'],
                    borderRadius: 5
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ y: {{ beginAtZero: true, max: 5 }} }}
            }}
        }});

        // Rating Distribution Pie Chart
        const ratingDistCtx = document.getElementById('ratingDistChart').getContext('2d');
        new Chart(ratingDistCtx, {{
            type: 'doughnut',
            data: {{
                labels: ['5-Star', '4-Star', '3-Star', 'Below 3'],
                datasets: [{{
                    data: [
                        dashboardData.rating_distribution['5-star'],
                        dashboardData.rating_distribution['4-star'],
                        dashboardData.rating_distribution['3-star'],
                        dashboardData.rating_distribution['below-3']
                    ],
                    backgroundColor: ['#4caf50', '#8bc34a', '#ffc107', '#f44336']
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ position: 'bottom' }} }}
            }}
        }});

        // Fulfillment Chart
        const fulfillmentCtx = document.getElementById('fulfillmentChart').getContext('2d');
        new Chart(fulfillmentCtx, {{
            type: 'doughnut',
            data: {{
                labels: Object.keys(dashboardData.fulfilled_by),
                datasets: [{{
                    data: Object.values(dashboardData.fulfilled_by),
                    backgroundColor: ['#667eea', '#764ba2', '#ff7675']
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ position: 'bottom' }} }}
            }}
        }});

        // Populate Tables
        function populateTable(tableId, data, showReviews = true) {{
            const tbody = document.getElementById(tableId);
            tbody.innerHTML = data.map(item => {{
                let row = `
                    <tr>
                        <td>${{item.Title}}</td>
                        <td>${{item.Price ? 'AED ' + item.Price.toFixed(2) : 'N/A'}}</td>
                        <td><span class="rating-badge">${{item.Rating}}</span></td>
                `;
                if (showReviews) {{
                    row += `<td>${{item.Reviews || 'N/A'}}</td>`;
                }}
                row += `</tr>`;
                return row;
            }}).join('');
        }}

        // Switch Tabs
        function switchTab(tabName) {{
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            document.getElementById(tabName).classList.add('active');
            event.target.classList.add('active');
        }}

        // Initialize Tables
        populateTable('topRatedTable', dashboardData.top_rated);
        populateTable('mostReviewedTable', dashboardData.most_reviewed);
        populateTable('bestValueTable', dashboardData.best_value, false);
    </script>
</body>
</html>
"""

    return html

def main():
    """Main function to generate dashboard"""

    # Accept an explicit file path as an override; otherwise use the most
    # recently modified week-stamped output file.
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = find_latest_excel_file(config.OUTPUT_DIR, config.OUTPUT_FILENAME)
        if not excel_file:
            print(f"❌ No output file found matching {config.OUTPUT_FILENAME} in {config.OUTPUT_DIR}")
            return False

    dashboard_file = os.path.join(config.OUTPUT_DIR, DASHBOARD_FILENAME)

    print(f"📊 Generating Dashboard...")
    print(f"   Reading: {excel_file}")

    # Read Excel data
    df = read_excel_data(excel_file)
    if df is None:
        print("❌ Failed to read Excel file")
        return False

    print(f"   Found {len(df)} products")

    # Calculate metrics
    print("   Calculating metrics...")
    metrics = calculate_metrics(df)
    if metrics is None:
        print("❌ Failed to calculate metrics")
        return False

    # Generate HTML
    print("   Generating HTML...")
    html = generate_html(metrics)

    # Save dashboard
    try:
        with open(dashboard_file, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"✅ Dashboard generated: {dashboard_file}")
        return True
    except Exception as e:
        print(f"❌ Error saving dashboard: {e}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
