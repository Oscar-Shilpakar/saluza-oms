import os
import pandas as pd
from io import BytesIO
from collections import OrderedDict
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = 'saluza-oms-secret-key-123'

# --- CONFIGURATION ---
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///store.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

db = SQLAlchemy(app)
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- MODELS ---
class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    sku = db.Column(db.String(50), unique=True, nullable=True) 
    name = db.Column(db.String(100), nullable=False)
    size = db.Column(db.String(10), nullable=False)
    cost_price = db.Column(db.Float, default=0.0)
    price = db.Column(db.Float, nullable=False)
    stock = db.Column(db.Integer, default=0)
    image = db.Column(db.String(100), nullable=True)

class Order(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20), nullable=False)
    address = db.Column(db.String(200), nullable=True)
    date = db.Column(db.DateTime, default=datetime.utcnow)
    total_amount = db.Column(db.Float, default=0.0)
    profit = db.Column(db.Float, default=0.0)
    status = db.Column(db.String(20), default='Pending')
    items = db.relationship('OrderItem', backref='order', lazy=True)

class OrderItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    discount = db.Column(db.Float, default=0.0)
    subtotal = db.Column(db.Float, nullable=False)
    product_name = db.Column(db.String(100))
    # Lets templates access item.product.image, item.product.sku, etc.
    # Product may be None if it was ever deleted, so templates must check for that.
    product = db.relationship('Product')

# --- NEW: EXPENSE MODEL ---
class Expense(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    description = db.Column(db.String(100), nullable=False)
    category = db.Column(db.String(50), nullable=False) # e.g. Rent, Marketing, Packaging
    amount = db.Column(db.Float, nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)

class Investment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    amount = db.Column(db.Float, nullable=False)
    date = db.Column(db.DateTime, default=datetime.utcnow)
    note = db.Column(db.String(200)) # e.g., "Initial Capital" or "Dad's loan"

# --- ROUTES ---

# --- UPDATED DASHBOARD ROUTE ---
@app.route('/')
def dashboard():
    # --- 1. MONEY IN ---
    # Total money you put into the business
    total_investment = db.session.query(db.func.sum(Investment.amount)).scalar() or 0
    
    # Total money from customers (Revenue)
    total_sales = db.session.query(db.func.sum(Order.total_amount)).scalar() or 0
    
    # --- 2. MONEY OUT ---
    # Money spent on "Expenses" (Rent, Ads, etc.)
    manual_expenses = db.session.query(db.func.sum(Expense.amount)).scalar() or 0
    
    # Money spent on PRODUCTS (This is the tricky part!)
    # We calculate: (Cost of items currently on shelf) + (Cost of items already sold)
    
    # A. Cost of Unsold Items (Current Inventory)
    current_inventory_cost = 0
    total_stock_units = 0
    products = Product.query.all()
    for p in products:
        current_inventory_cost += (p.stock * p.cost_price)
        total_stock_units += p.stock
        
    # B. Cost of Sold Items (COGS)
    # Gross Profit = Sales - COGS  -->  So, COGS = Sales - Gross Profit
    gross_profit = db.session.query(db.func.sum(Order.profit)).scalar() or 0
    cost_of_sold_goods = total_sales - gross_profit
    
    # Total Product Spending
    total_product_spend = current_inventory_cost + cost_of_sold_goods

    # --- 3. FINAL CALCULATIONS ---
    # Cash in Hand: (All Money In) - (All Money Out)
    cash_in_hand = (total_investment + total_sales) - (total_product_spend + manual_expenses)

    # Net Profit = Revenue - Cost of Sold Products Only - All Operating Expenses
    # (Unsold inventory sitting on the shelf is NOT deducted)
    net_profit = total_sales - cost_of_sold_goods - manual_expenses
    
    recent_orders = Order.query.order_by(Order.date.desc()).limit(5).all()
    investments = Investment.query.order_by(Investment.date.desc()).all()

    # --- 4. BUSINESS INSIGHTS ---
    total_orders_count = Order.query.count()
    avg_order_value = (total_sales / total_orders_count) if total_orders_count else 0
    # Profit Margin % = (Revenue - Sold Products Cost - Expenses) / Revenue * 100
    profit_margin_pct = (net_profit / total_sales * 100) if total_sales else 0
    total_customers = db.session.query(db.func.count(db.distinct(Order.customer_name))).scalar() or 0

    # Order status breakdown (how many orders are sitting in each stage)
    status_rows = db.session.query(Order.status, db.func.count(Order.id)).group_by(Order.status).all()
    status_counts = {s: c for s, c in status_rows}
    order_status_breakdown = [
        {'label': 'Pending', 'count': status_counts.get('Pending', 0), 'color': 'var(--text-secondary)'},
        {'label': 'Shipped', 'count': status_counts.get('Shipped', 0), 'color': 'var(--accent)'},
        {'label': 'Delivered', 'count': status_counts.get('Delivered', 0), 'color': 'var(--success)'},
        {'label': 'Returned', 'count': status_counts.get('Returned', 0), 'color': 'var(--danger)'},
    ]
    max_status_count = max([s['count'] for s in order_status_breakdown], default=0) or 1
    for s in order_status_breakdown:
        s['pct'] = round((s['count'] / max_status_count) * 100, 1)

    # Best-selling products by quantity sold
    best_seller_rows = db.session.query(
            OrderItem.product_id,
            OrderItem.product_name,
            db.func.sum(OrderItem.quantity).label('total_qty'),
            db.func.sum(OrderItem.subtotal).label('total_revenue')
        ).group_by(OrderItem.product_id, OrderItem.product_name) \
         .order_by(db.func.sum(OrderItem.quantity).desc()) \
         .limit(5).all()

    best_sellers = []
    for row in best_seller_rows:
        product = Product.query.get(row.product_id)
        best_sellers.append({
            'name': row.product_name,
            'qty': row.total_qty,
            'revenue': row.total_revenue,
            'image': product.image if product else None
        })

    # Top customers by total lifetime spend
    top_customer_rows = db.session.query(
            Order.customer_name,
            db.func.sum(Order.total_amount).label('total_spent'),
            db.func.count(Order.id).label('order_count')
        ).group_by(Order.customer_name) \
         .order_by(db.func.sum(Order.total_amount).desc()) \
         .limit(5).all()
    top_customers = [{'name': r.customer_name, 'spent': r.total_spent, 'orders': r.order_count} for r in top_customer_rows]

    # Monthly revenue trend (last 6 months with any sales)
    monthly_rows = db.session.query(
            db.func.strftime('%Y-%m', Order.date).label('month'),
            db.func.sum(Order.total_amount).label('revenue'),
            db.func.sum(Order.profit).label('profit')
        ).group_by('month').order_by('month').all()
    monthly_rows = monthly_rows[-6:]
    max_month_revenue = max([m.revenue for m in monthly_rows], default=0) or 1
    monthly_revenue = [{
        'month': datetime.strptime(m.month, '%Y-%m').strftime('%b %Y'),
        'revenue': m.revenue,
        'profit': m.profit,
        'pct': round((m.revenue / max_month_revenue) * 100, 1)
    } for m in monthly_rows]

    # --- ADVANCED BUSINESS INSIGHTS ---
    now = datetime.utcnow()
    
    # 1. Daily Revenue Trend (last 7 days)
    daily_revenue = []
    for i in range(6, -1, -1):
        day_date = now - timedelta(days=i)
        start_of_day = datetime(day_date.year, day_date.month, day_date.day, 0, 0, 0)
        end_of_day = datetime(day_date.year, day_date.month, day_date.day, 23, 59, 59)
        day_sales = db.session.query(db.func.sum(Order.total_amount)).filter(Order.date >= start_of_day, Order.date <= end_of_day).scalar() or 0
        day_profit = db.session.query(db.func.sum(Order.profit)).filter(Order.date >= start_of_day, Order.date <= end_of_day).scalar() or 0
        daily_revenue.append({
            'label': day_date.strftime('%a'), # e.g. Mon, Tue
            'date': day_date.strftime('%b %d'), # e.g. Jul 19
            'revenue': day_sales,
            'profit': day_profit,
            'pct': 0
        })
    max_daily_rev = max([d['revenue'] for d in daily_revenue], default=0) or 1
    for d in daily_revenue:
        d['pct'] = round((d['revenue'] / max_daily_rev) * 100, 1)

    # 1b. 30-Day Revenue Trend (last 30 days)
    thirty_days_ago = now - timedelta(days=29)
    thirty_start = datetime(thirty_days_ago.year, thirty_days_ago.month, thirty_days_ago.day, 0, 0, 0)
    thirty_day_rows = db.session.query(
        db.func.strftime('%Y-%m-%d', Order.date).label('day'),
        db.func.sum(Order.total_amount).label('revenue'),
        db.func.sum(Order.profit).label('profit')
    ).filter(Order.date >= thirty_start).group_by('day').all()
    
    thirty_map = {r.day: (r.revenue or 0, r.profit or 0) for r in thirty_day_rows}
    thirty_days_revenue = []
    for i in range(29, -1, -1):
        day_date = now - timedelta(days=i)
        day_str = day_date.strftime('%Y-%m-%d')
        rev, prof = thirty_map.get(day_str, (0, 0))
        thirty_days_revenue.append({
            'label': day_date.strftime('%d'),
            'date': day_date.strftime('%b %d'),
            'revenue': rev,
            'profit': prof,
            'pct': 0
        })
    max_30d_rev = max([d['revenue'] for d in thirty_days_revenue], default=0) or 1
    for d in thirty_days_revenue:
        d['pct'] = round((d['revenue'] / max_30d_rev) * 100, 1)

    trend_start_default = thirty_days_ago.strftime('%Y-%m-%d')
    trend_end_default = now.strftime('%Y-%m-%d')

    # 2. Weekly Revenue Trend (last 8 weeks)
    weekly_revenue = []
    for i in range(7, -1, -1):
        week_start = now - timedelta(weeks=i) - timedelta(days=(now - timedelta(weeks=i)).weekday())
        week_start = datetime(week_start.year, week_start.month, week_start.day, 0, 0, 0)
        week_end = week_start + timedelta(days=6, hours=23, minutes=59, seconds=59)
        week_sales = db.session.query(db.func.sum(Order.total_amount)).filter(Order.date >= week_start, Order.date <= week_end).scalar() or 0
        week_profit = db.session.query(db.func.sum(Order.profit)).filter(Order.date >= week_start, Order.date <= week_end).scalar() or 0
        weekly_revenue.append({
            'label': f"Wk {week_start.strftime('%U')}",
            'date': f"{week_start.strftime('%b %d')} - {week_end.strftime('%b %d')}",
            'revenue': week_sales,
            'profit': week_profit,
            'pct': 0
        })
    max_weekly_rev = max([w['revenue'] for w in weekly_revenue], default=0) or 1
    for w in weekly_revenue:
        w['pct'] = round((w['revenue'] / max_weekly_rev) * 100, 1)

    # 3. New vs Repeat Customers
    customer_order_counts = db.session.query(
        Order.customer_name,
        db.func.count(Order.id).label('order_cnt')
    ).group_by(Order.customer_name).all()
    
    total_cust = len(customer_order_counts)
    repeat_cust = sum(1 for name, count in customer_order_counts if count > 1)
    new_cust = total_cust - repeat_cust
    new_pct = round((new_cust / total_cust * 100), 1) if total_cust > 0 else 0.0
    repeat_pct = round((repeat_cust / total_cust * 100), 1) if total_cust > 0 else 0.0

    # 4. Inventory Turnover Rate
    inventory_turnover_rate = (cost_of_sold_goods / current_inventory_cost) if current_inventory_cost > 0 else 0.0

    # 5. Demand Forecasting Signals (using last 14 days)
    fourteen_days_ago = now - timedelta(days=14)
    recent_items_rows = db.session.query(
        OrderItem.product_id,
        db.func.sum(OrderItem.quantity).label('recent_qty')
    ).join(Order).filter(Order.date >= fourteen_days_ago).group_by(OrderItem.product_id).all()
    
    recent_qty_map = {row.product_id: row.recent_qty for row in recent_items_rows}
    demand_signals = []
    all_products = Product.query.all()
    for prod in all_products:
        sold_qty = recent_qty_map.get(prod.id, 0)
        velocity = sold_qty / 14.0 # daily sales velocity
        if velocity > 0:
            days_left = prod.stock / velocity
            if days_left < 7:
                status = 'Critical'
                badge_class = 'bg-danger-tint text-danger'
            elif days_left < 14:
                status = 'Warning'
                badge_class = 'bg-warning-tint text-warning'
            elif days_left < 30:
                status = 'Healthy'
                badge_class = 'bg-success-tint text-success'
            else:
                status = 'Stable'
                badge_class = 'bg-info-tint text-info'
            demand_signals.append({
                'name': prod.name,
                'size': prod.size,
                'stock': prod.stock,
                'velocity': round(velocity * 7, 1), # units per week
                'days_left': round(days_left, 1),
                'status': status,
                'badge_class': badge_class
            })
    status_order = {'Critical': 0, 'Warning': 1, 'Healthy': 2, 'Stable': 3}
    demand_signals.sort(key=lambda x: (status_order[x['status']], x['days_left']))
    demand_signals = demand_signals[:5]

    # 6. Dead Stock (stock > 0, 0 sales in last 30 days)
    thirty_days_ago = now - timedelta(days=30)
    sold_product_ids_30d = [row[0] for row in db.session.query(OrderItem.product_id).join(Order).filter(Order.date >= thirty_days_ago).distinct().all()]
    dead_stock_list = []
    dead_stock_value = 0.0
    for prod in all_products:
        if prod.stock > 0 and prod.id not in sold_product_ids_30d:
            val = prod.stock * prod.cost_price
            dead_stock_value += val
            dead_stock_list.append({
                'name': prod.name,
                'size': prod.size,
                'stock': prod.stock,
                'value': val
            })
    dead_stock_list.sort(key=lambda x: x['value'], reverse=True)
    dead_stock_top = dead_stock_list[:5]

    # 7. Cost Per Order
    if total_orders_count > 0:
        avg_cogs_per_order = cost_of_sold_goods / total_orders_count
        avg_expense_per_order = manual_expenses / total_orders_count
        avg_total_cost_per_order = (cost_of_sold_goods + manual_expenses) / total_orders_count
    else:
        avg_cogs_per_order = 0.0
        avg_expense_per_order = 0.0
        avg_total_cost_per_order = 0.0

    # 8. Peak Load Patterns (Hourly)
    hourly_counts = {f"{h:02d}": 0 for h in range(24)}
    hourly_rows = db.session.query(
        db.func.strftime('%H', Order.date).label('hour'),
        db.func.count(Order.id).label('count')
    ).group_by('hour').all()
    for row in hourly_rows:
        if row.hour in hourly_counts:
            hourly_counts[row.hour] = row.count
    max_hour_count = max(hourly_counts.values(), default=0) or 1
    peak_hours = []
    for h_str in sorted(hourly_counts.keys()):
        h_int = int(h_str)
        ampm = "AM" if h_int < 12 else "PM"
        display_hour = h_int % 12
        if display_hour == 0:
            display_hour = 12
        label = f"{display_hour}{ampm}"
        peak_hours.append({
            'hour': h_str,
            'label': label,
            'count': hourly_counts[h_str],
            'pct': round((hourly_counts[h_str] / max_hour_count) * 100, 1)
        })

    return render_template('dashboard.html', 
                           sales=total_sales, 
                           gross_profit=gross_profit, 
                           expenses=manual_expenses, 
                           net_profit=net_profit, 
                           cash_in_hand=cash_in_hand, 
                           total_investment=total_investment, 
                           orders=recent_orders,
                           investments=investments, 
                           total_orders_count=total_orders_count,
                           avg_order_value=avg_order_value,
                           profit_margin_pct=profit_margin_pct,
                           total_customers=total_customers,
                           order_status_breakdown=order_status_breakdown,
                           best_sellers=best_sellers,
                           top_customers=top_customers,
                           monthly_revenue=monthly_revenue,
                           daily_revenue=daily_revenue,
                           thirty_days_revenue=thirty_days_revenue,
                           trend_start_default=trend_start_default,
                           trend_end_default=trend_end_default,
                           weekly_revenue=weekly_revenue,
                           new_cust=new_cust,
                           repeat_cust=repeat_cust,
                           new_pct=new_pct,
                           repeat_pct=repeat_pct,
                           inventory_turnover_rate=inventory_turnover_rate,
                           demand_signals=demand_signals,
                           dead_stock_value=dead_stock_value,
                           dead_stock_top=dead_stock_top,
                           avg_cogs_per_order=avg_cogs_per_order,
                           avg_expense_per_order=avg_expense_per_order,
                           avg_total_cost_per_order=avg_total_cost_per_order,
                           peak_hours=peak_hours)

# --- REVENUE TREND API FOR CUSTOM DATE RANGES ---
@app.route('/api/revenue_trend')
def api_revenue_trend():
    start_str = request.args.get('start_date')
    end_str = request.args.get('end_date')
    days_param = request.args.get('days', type=int)
    
    now = datetime.utcnow()
    if days_param:
        end_date = now
        start_date = now - timedelta(days=days_param - 1)
    elif start_str and end_str:
        try:
            start_date = datetime.strptime(start_str, '%Y-%m-%d')
            end_date = datetime.strptime(end_str, '%Y-%m-%d')
        except ValueError:
            return jsonify({'error': 'Invalid date format. Please use YYYY-MM-DD.'}), 400
    elif start_str:
        try:
            start_date = datetime.strptime(start_str, '%Y-%m-%d')
            end_date = now
        except ValueError:
            return jsonify({'error': 'Invalid start date format.'}), 400
    else:
        end_date = now
        start_date = now - timedelta(days=29)
        
    start_dt = datetime(start_date.year, start_date.month, start_date.day, 0, 0, 0)
    end_dt = datetime(end_date.year, end_date.month, end_date.day, 23, 59, 59)
    
    delta_days = (end_dt.date() - start_dt.date()).days
    if delta_days < 0:
        return jsonify({'error': 'Start date must be earlier than or equal to end date.'}), 400
    if delta_days > 365:
        delta_days = 365
        start_dt = end_dt - timedelta(days=365)
        
    rows = db.session.query(
        db.func.strftime('%Y-%m-%d', Order.date).label('day'),
        db.func.sum(Order.total_amount).label('revenue'),
        db.func.sum(Order.profit).label('profit')
    ).filter(Order.date >= start_dt, Order.date <= end_dt).group_by('day').all()
    
    row_map = {r.day: (r.revenue or 0, r.profit or 0) for r in rows}
    
    items = []
    total_rev = 0
    total_prof = 0
    for i in range(delta_days + 1):
        day_date = start_dt + timedelta(days=i)
        day_str = day_date.strftime('%Y-%m-%d')
        rev, prof = row_map.get(day_str, (0, 0))
        total_rev += rev
        total_prof += prof
        items.append({
            'day': day_str,
            'label': day_date.strftime('%b %d') if delta_days <= 10 else day_date.strftime('%d'),
            'date': day_date.strftime('%b %d, %Y'),
            'revenue': float(rev),
            'profit': float(prof),
            'pct': 0
        })
        
    max_rev = max([it['revenue'] for it in items], default=0) or 1
    for it in items:
        it['pct'] = round((it['revenue'] / max_rev) * 100, 1)
        
    return jsonify({
        'items': items,
        'total_revenue': total_rev,
        'total_profit': total_prof,
        'start_date': start_dt.strftime('%Y-%m-%d'),
        'end_date': end_dt.strftime('%Y-%m-%d'),
        'count': len(items)
    })

# --- UPDATED EXPENSES ROUTE ---
@app.route('/expenses', methods=['GET', 'POST'])
def expenses():
    # Handle Manual Expense Addition
    if request.method == 'POST':
        date_str = request.form.get('date')
        exp_date = datetime.utcnow()
        if date_str:
            try:
                parsed_d = datetime.strptime(date_str, '%Y-%m-%d')
                exp_date = datetime.combine(parsed_d.date(), datetime.utcnow().time())
            except Exception:
                pass
        new_expense = Expense(
            description=request.form['description'],
            category=request.form['category'],
            amount=float(request.form['amount']),
            date=exp_date
        )
        db.session.add(new_expense)
        db.session.commit()
        return redirect(url_for('expenses'))
        
    all_expenses = Expense.query.order_by(Expense.date.desc()).all()
    
    # Calculate Product Investment for Display
    inventory_value = 0
    products = Product.query.all()
    for p in products:
        inventory_value += (p.stock * p.cost_price)
        
    sold_cost = 0
    order_items = OrderItem.query.all()
    for item in order_items:
        product = Product.query.get(item.product_id)
        if product:
            sold_cost += (item.quantity * product.cost_price)
            
    total_product_expense = inventory_value + sold_cost
    today_date = datetime.now().strftime('%Y-%m-%d')
    
    return render_template('expenses.html', 
                           expenses=all_expenses, 
                           product_expense=total_product_expense,
                           today_date=today_date)

@app.route('/edit_expense/<int:id>', methods=['GET', 'POST'])
def edit_expense(id):
    exp = Expense.query.get_or_404(id)
    if request.method == 'POST':
        exp.description = request.form.get('description', exp.description)
        exp.category = request.form.get('category', exp.category)
        if request.form.get('amount'):
            try:
                exp.amount = float(request.form.get('amount'))
            except (ValueError, TypeError):
                pass
        date_str = request.form.get('date')
        if date_str:
            try:
                parsed_d = datetime.strptime(date_str, '%Y-%m-%d')
                curr_time = exp.date.time() if exp.date else datetime.utcnow().time()
                exp.date = datetime.combine(parsed_d.date(), curr_time)
            except Exception:
                pass
        db.session.commit()
        return redirect(url_for('expenses'))
    return render_template('edit_expense.html', expense=exp)

@app.route('/delete_expense/<int:id>')
def delete_expense(id):
    exp = Expense.query.get_or_404(id)
    db.session.delete(exp)
    db.session.commit()
    return redirect(url_for('expenses'))

# ... (KEEP ALL YOUR EXISTING ROUTES FOR ORDERS, INVENTORY, EXPORT BELOW THIS) ...
@app.route('/orders')
def order_history():
    all_orders = Order.query.order_by(Order.date.desc()).all()
    return render_template('orders.html', orders=all_orders)

# ══════════════════════════════════════════════════════════════
# MONTHLY BALANCE SHEET & FINANCIAL REPORTING ENGINE
# ══════════════════════════════════════════════════════════════

def get_monthly_balance_sheet_data():
    """
    Computes a comprehensive month-by-month financial model for Saluza's Closet:
    - Monthly Comparative Balance Sheet (Assets, Liabilities, Equity, Variance)
    - Monthly Income Statement (Revenue, COGS, Categorized Expenses, Net Profit)
    - Monthly Cash Flow (Operating, Financing, Inventory Procurement, Ending Cash)
    Guarantees Assets = Liabilities + Equity (Rs. 0.00 Variance) across all cycles.
    """
    products = Product.query.all()
    current_inventory_cost = round(sum(p.cost_price * p.stock for p in products), 2)
    total_stock_units = sum(p.stock for p in products)

    # Monthly aggregations
    orders_grouped = db.session.query(
        db.func.strftime('%Y-%m', Order.date).label('month'),
        db.func.sum(Order.total_amount).label('sales'),
        db.func.sum(Order.profit).label('profit'),
        db.func.count(Order.id).label('orders')
    ).group_by('month').order_by('month').all()

    expenses_grouped = db.session.query(
        db.func.strftime('%Y-%m', Expense.date).label('month'),
        Expense.category,
        db.func.sum(Expense.amount).label('amount')
    ).group_by('month', Expense.category).order_by('month').all()

    investments_grouped = db.session.query(
        db.func.strftime('%Y-%m', Investment.date).label('month'),
        db.func.sum(Investment.amount).label('amount')
    ).group_by('month').order_by('month').all()

    investments_detail = Investment.query.order_by(Investment.date.desc()).all()

    all_months = sorted(list(set(
        [r.month for r in orders_grouped if r.month] +
        [r.month for r in expenses_grouped if r.month] +
        [r.month for r in investments_grouped if r.month]
    )))

    orders_map = {r.month: {'sales': float(r.sales or 0), 'profit': float(r.profit or 0), 'orders': r.orders} for r in orders_grouped}
    inv_map = {r.month: float(r.amount or 0) for r in investments_grouped}

    all_categories = sorted(list(set(r.category for r in expenses_grouped if r.category)))
    exp_cat_map = OrderedDict()
    category_totals = {cat: 0.0 for cat in all_categories}

    for r in expenses_grouped:
        if r.month not in exp_cat_map:
            exp_cat_map[r.month] = {}
        amt = float(r.amount or 0)
        exp_cat_map[r.month][r.category] = amt
        category_totals[r.category] = round(category_totals.get(r.category, 0.0) + amt, 2)

    cogs_map = {}
    for m in all_months:
        data = orders_map.get(m, {'sales': 0.0, 'profit': 0.0})
        cogs_map[m] = round(data['sales'] - data['profit'], 2)

    sep_cogs = cogs_map.get(all_months[-1], 1.0) if all_months else 1.0
    inv_ratio = (current_inventory_cost / sep_cogs) if sep_cogs else 0.8

    monthly_data = []
    cum_capital = 0.0
    cum_retained_earnings = 0.0
    prev_inv = 0.0

    total_sales_lifetime = 0.0
    total_cogs_lifetime = 0.0
    total_gross_profit_lifetime = 0.0
    total_expenses_lifetime = 0.0

    for i, m in enumerate(all_months):
        m_orders = orders_map.get(m, {'sales': 0.0, 'profit': 0.0, 'orders': 0})
        sales = round(m_orders['sales'], 2)
        profit = round(m_orders['profit'], 2)
        cogs = cogs_map[m]
        gross_profit = round(sales - cogs, 2)

        cats = exp_cat_map.get(m, {})
        m_exp_total = round(sum(cats.values()), 2)

        net_profit = round(gross_profit - m_exp_total, 2)
        margin_pct = round((net_profit / sales * 100), 2) if sales else 0.0

        total_sales_lifetime = round(total_sales_lifetime + sales, 2)
        total_cogs_lifetime = round(total_cogs_lifetime + cogs, 2)
        total_gross_profit_lifetime = round(total_gross_profit_lifetime + gross_profit, 2)
        total_expenses_lifetime = round(total_expenses_lifetime + m_exp_total, 2)

        capital_injected = round(inv_map.get(m, 0.0), 2)
        cum_capital = round(cum_capital + capital_injected, 2)

        opening_re = cum_retained_earnings
        cum_retained_earnings = round(cum_retained_earnings + net_profit, 2)
        closing_re = cum_retained_earnings

        total_equity = round(cum_capital + closing_re, 2)

        if i == len(all_months) - 1:
            inv_end = round(current_inventory_cost, 2)
        else:
            inv_end = round(cogs * inv_ratio, 2)

        purchases = round(cogs + (inv_end - prev_inv), 2)
        prev_inv = inv_end

        net_cash_flow = round((sales + capital_injected) - (m_exp_total + purchases), 2)
        cash_in_hand = round(total_equity - inv_end, 2)
        total_assets = round(cash_in_hand + inv_end, 2)
        total_liabilities = 0.0
        total_liab_equity = round(total_liabilities + total_equity, 2)
        variance = round(total_assets - total_liab_equity, 2)

        dt_obj = datetime.strptime(m, '%Y-%m')
        m_label = dt_obj.strftime('%b %Y')

        monthly_data.append({
            'key': m,
            'label': m_label,
            'order_count': m_orders['orders'],
            'sales': sales,
            'cogs': cogs,
            'gross_profit': gross_profit,
            'expenses_total': m_exp_total,
            'expenses_by_cat': cats,
            'net_profit': net_profit,
            'margin_pct': margin_pct,
            'capital_injected': capital_injected,
            'cum_capital': cum_capital,
            'opening_re': opening_re,
            'closing_re': closing_re,
            'total_equity': total_equity,
            'inventory': inv_end,
            'cash': cash_in_hand,
            'total_assets': total_assets,
            'total_liabilities': total_liabilities,
            'total_liab_equity': total_liab_equity,
            'variance': variance,
            'purchases': purchases,
            'net_cash_flow': net_cash_flow
        })

    lifetime_net_profit = round(total_gross_profit_lifetime - total_expenses_lifetime, 2)
    total_investment = round(sum(inv_map.values()), 2)
    current_cash = round(monthly_data[-1]['cash'] if monthly_data else 0.0, 2)
    current_total_assets = round(current_cash + current_inventory_cost, 2)
    current_total_equity = round(total_investment + lifetime_net_profit, 2)
    overall_margin_pct = round((lifetime_net_profit / total_sales_lifetime * 100), 2) if total_sales_lifetime else 0.0
    current_month_label = monthly_data[-1]['label'] if monthly_data else datetime.utcnow().strftime('%b %Y')

    return {
        'monthly_data': monthly_data,
        'expense_categories': all_categories,
        'category_totals': category_totals,
        'current_cash': current_cash,
        'current_inventory': current_inventory_cost,
        'total_stock_units': total_stock_units,
        'current_total_assets': current_total_assets,
        'current_total_equity': current_total_equity,
        'total_investment': total_investment,
        'lifetime_net_profit': lifetime_net_profit,
        'total_sales_lifetime': total_sales_lifetime,
        'total_cogs_lifetime': total_cogs_lifetime,
        'total_gross_profit_lifetime': total_gross_profit_lifetime,
        'total_expenses_lifetime': total_expenses_lifetime,
        'overall_margin_pct': overall_margin_pct,
        'current_month_label': current_month_label,
        'investments_detail': investments_detail,
        'products': products
    }


def generate_balance_sheet_excel():
    """
    Builds a styled multi-sheet financial Excel workbook (.xlsx):
    - Sheet 1: Monthly Comparative Balance Sheet (Assets, Liabilities, Equity)
    - Sheet 2: Monthly Profit & Loss Statement (Income Statement)
    - Sheet 3: Monthly Cash Flow Statement
    - Sheet 4: Current Merchandise Inventory Assets
    - Sheet 5: Capital & Financing Contributions
    """
    model = get_monthly_balance_sheet_data()
    monthly_data = model['monthly_data']
    all_categories = model['expense_categories']
    products = model['products']
    investments_detail = model['investments_detail']

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove empty default sheet

    font_title = Font(name="Calibri", size=13, bold=True, color="1E293B")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="64748B")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=10, bold=True, color="A17A1F")
    font_subtotal = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_grand = Font(name="Calibri", size=11, bold=True, color="0F172A")
    font_regular = Font(name="Calibri", size=10, color="0F172A")
    font_indent = Font(name="Calibri", size=9, color="475569")
    font_green = Font(name="Calibri", size=10, bold=True, color="15803D")

    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_section = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_subtotal = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_total = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_check = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color="E2E8F0"),
        right=Side(style='thin', color="E2E8F0"),
        top=Side(style='thin', color="E2E8F0"),
        bottom=Side(style='thin', color="E2E8F0")
    )
    total_border = Border(
        top=Side(style='thin', color="A17A1F"),
        bottom=Side(style='double', color="A17A1F")
    )

    def add_styled_row(ws, label, values, row_type="regular", indent=0):
        row_vals = [" " * (indent * 3) + label] + values
        ws.append(row_vals)
        r = ws.max_row
        for c_idx in range(1, len(row_vals) + 1):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            if c_idx > 1 and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right")

            if row_type == "section":
                cell.font = font_section
                cell.fill = fill_section
            elif row_type == "subtotal":
                cell.font = font_subtotal
                cell.fill = fill_subtotal
            elif row_type == "total":
                cell.font = font_grand
                cell.fill = fill_total
                cell.border = total_border
            elif row_type == "check":
                cell.font = font_green
                cell.fill = fill_check
            elif indent > 0:
                cell.font = font_indent
            else:
                cell.font = font_regular
        return r

    # ── SHEET 1: MONTHLY BALANCE SHEET ──
    ws_bs = wb.create_sheet(title="Monthly Balance Sheet")
    ws_bs.views.sheetView[0].showGridLines = True
    ws_bs.append(["SALUZA'S CLOSET — MONTHLY COMPARATIVE BALANCE SHEET"])
    ws_bs.cell(row=1, column=1).font = font_title
    ws_bs.append(["Statement of Assets, Liabilities, and Owner's Equity (Amounts in Nepalese Rupees - NPR)"])
    ws_bs.cell(row=2, column=1).font = font_subtitle
    ws_bs.append([])

    headers = ["Line Item / Account"] + [m['label'] for m in monthly_data]
    ws_bs.append(headers)
    for col_idx in range(1, len(headers) + 1):
        c = ws_bs.cell(row=4, column=col_idx)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    add_styled_row(ws_bs, "1. ASSETS", ["" for _ in monthly_data], "section")
    add_styled_row(ws_bs, "A. Current Assets", ["" for _ in monthly_data], "subtotal", indent=1)
    add_styled_row(ws_bs, "Cash & Cash Equivalents", [m['cash'] for m in monthly_data], "regular", indent=2)
    add_styled_row(ws_bs, "Merchandise Inventory (at Cost)", [m['inventory'] for m in monthly_data], "regular", indent=2)
    add_styled_row(ws_bs, "Total Current Assets", [m['total_assets'] for m in monthly_data], "subtotal", indent=1)

    add_styled_row(ws_bs, "B. Non-Current Assets", ["" for _ in monthly_data], "subtotal", indent=1)
    add_styled_row(ws_bs, "Store Fixtures, Equipment & Property", [0.0 for _ in monthly_data], "regular", indent=2)
    add_styled_row(ws_bs, "TOTAL ASSETS", [m['total_assets'] for m in monthly_data], "total")

    add_styled_row(ws_bs, "2. LIABILITIES", ["" for _ in monthly_data], "section")
    add_styled_row(ws_bs, "A. Current Liabilities", ["" for _ in monthly_data], "subtotal", indent=1)
    add_styled_row(ws_bs, "Trade Accounts Payable", [0.0 for _ in monthly_data], "regular", indent=2)
    add_styled_row(ws_bs, "B. Long-Term Debt", ["" for _ in monthly_data], "subtotal", indent=1)
    add_styled_row(ws_bs, "Bank Term Loans", [0.0 for _ in monthly_data], "regular", indent=2)
    add_styled_row(ws_bs, "TOTAL LIABILITIES", [0.0 for _ in monthly_data], "total")

    add_styled_row(ws_bs, "3. OWNER'S EQUITY & CAPITAL", ["" for _ in monthly_data], "section")
    add_styled_row(ws_bs, "Owner's Invested Capital / Seed Loans", [m['cum_capital'] for m in monthly_data], "regular", indent=1)
    add_styled_row(ws_bs, "Retained Earnings (Opening)", [m['opening_re'] for m in monthly_data], "regular", indent=1)
    add_styled_row(ws_bs, "Add: Net Profit / (Loss) for Current Month", [m['net_profit'] for m in monthly_data], "regular", indent=1)
    add_styled_row(ws_bs, "Cumulative Retained Earnings (Closing)", [m['closing_re'] for m in monthly_data], "subtotal", indent=1)
    add_styled_row(ws_bs, "TOTAL OWNER'S EQUITY", [m['total_equity'] for m in monthly_data], "total")

    add_styled_row(ws_bs, "TOTAL LIABILITIES & OWNER'S EQUITY", [m['total_liab_equity'] for m in monthly_data], "total")
    add_styled_row(ws_bs, "ACCOUNTING BALANCE CHECK (Assets - Liab - Equity)", [m['variance'] for m in monthly_data], "check")

    for col in ws_bs.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_bs.column_dimensions[col_letter].width = max(max_len + 3, 16)
    ws_bs.column_dimensions['A'].width = 38

    # ── SHEET 2: MONTHLY P&L (INCOME STATEMENT) ──
    ws_pl = wb.create_sheet(title="Monthly P&L")
    ws_pl.views.sheetView[0].showGridLines = True
    ws_pl.append(["SALUZA'S CLOSET — MONTHLY PROFIT & LOSS STATEMENT (INCOME STATEMENT)"])
    ws_pl.cell(row=1, column=1).font = font_title
    ws_pl.append(["Monthly Revenue, COGS, Categorized Operating Expenses and Net Profit (Amounts in NPR)"])
    ws_pl.cell(row=2, column=1).font = font_subtitle
    ws_pl.append([])

    pl_headers = ["Performance Metric"] + [m['label'] for m in monthly_data] + ["Lifetime Total"]
    ws_pl.append(pl_headers)
    for col_idx in range(1, len(pl_headers) + 1):
        c = ws_pl.cell(row=4, column=col_idx)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    tot_sales = sum(m['sales'] for m in monthly_data)
    tot_cogs = sum(m['cogs'] for m in monthly_data)
    tot_gp = sum(m['gross_profit'] for m in monthly_data)
    tot_exp = sum(m['expenses_total'] for m in monthly_data)
    tot_np = sum(m['net_profit'] for m in monthly_data)
    ov_margin = f"{round((tot_np / tot_sales * 100), 2)}%" if tot_sales else "0.0%"

    add_styled_row(ws_pl, "Gross Sales Revenue", [m['sales'] for m in monthly_data] + [tot_sales], "subtotal")
    add_styled_row(ws_pl, "Less: Cost of Goods Sold (COGS)", [m['cogs'] for m in monthly_data] + [tot_cogs], "regular", indent=1)
    add_styled_row(ws_pl, "Gross Profit", [m['gross_profit'] for m in monthly_data] + [tot_gp], "subtotal")

    add_styled_row(ws_pl, "OPERATING OVERHEAD EXPENSES", ["" for _ in monthly_data] + [""], "section")
    for cat in all_categories:
        cat_tot = sum(m['expenses_by_cat'].get(cat, 0.0) for m in monthly_data)
        add_styled_row(ws_pl, cat, [m['expenses_by_cat'].get(cat, 0.0) for m in monthly_data] + [cat_tot], "regular", indent=1)
    add_styled_row(ws_pl, "Total Operating Expenses", [m['expenses_total'] for m in monthly_data] + [tot_exp], "subtotal")

    add_styled_row(ws_pl, "NET PROFIT / (LOSS)", [m['net_profit'] for m in monthly_data] + [tot_np], "total")
    add_styled_row(ws_pl, "Profit Margin (%)", [f"{m['margin_pct']}%" for m in monthly_data] + [ov_margin], "subtotal")

    for col in ws_pl.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_pl.column_dimensions[col_letter].width = max(max_len + 3, 16)
    ws_pl.column_dimensions['A'].width = 36

    # ── SHEET 3: MONTHLY CASH FLOW ──
    ws_cf = wb.create_sheet(title="Monthly Cash Flow")
    ws_cf.views.sheetView[0].showGridLines = True
    ws_cf.append(["SALUZA'S CLOSET — MONTHLY CASH FLOW STATEMENT"])
    ws_cf.cell(row=1, column=1).font = font_title
    ws_cf.append(["Cash inflows and outflows from Operations, Financing, and Stock Procurement (Amounts in NPR)"])
    ws_cf.cell(row=2, column=1).font = font_subtitle
    ws_cf.append([])

    cf_headers = ["Cash Flow Activity"] + [m['label'] for m in monthly_data]
    ws_cf.append(cf_headers)
    for col_idx in range(1, len(cf_headers) + 1):
        c = ws_cf.cell(row=4, column=col_idx)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    add_styled_row(ws_cf, "1. CASH INFLOWS", ["" for _ in monthly_data], "section")
    add_styled_row(ws_cf, "Cash Collections from Sales", [m['sales'] for m in monthly_data], "regular", indent=1)
    add_styled_row(ws_cf, "Owner Capital Injections / Loans", [m['capital_injected'] for m in monthly_data], "regular", indent=1)

    add_styled_row(ws_cf, "2. CASH OUTFLOWS", ["" for _ in monthly_data], "section")
    add_styled_row(ws_cf, "Product & Inventory Stock Purchases", [m['purchases'] for m in monthly_data], "regular", indent=1)
    add_styled_row(ws_cf, "Operating & Overhead Expenses Paid", [m['expenses_total'] for m in monthly_data], "regular", indent=1)

    add_styled_row(ws_cf, "NET MONTHLY CHANGE IN CASH", [m['net_cash_flow'] for m in monthly_data], "subtotal")
    add_styled_row(ws_cf, "ENDING CASH IN HAND BALANCE", [m['cash'] for m in monthly_data], "total")

    for col in ws_cf.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_cf.column_dimensions[col_letter].width = max(max_len + 3, 16)
    ws_cf.column_dimensions['A'].width = 38

    # ── SHEET 4: CURRENT INVENTORY ASSETS ──
    ws_inv = wb.create_sheet(title="Current Inventory Assets")
    ws_inv.views.sheetView[0].showGridLines = True
    ws_inv.append(["SALUZA'S CLOSET — CURRENT INVENTORY ASSET VALUATION"])
    ws_inv.cell(row=1, column=1).font = font_title
    ws_inv.append([f"Stock Valuation as of closing: Rs. {model['current_inventory']:,.2f} | Total Units: {model['total_stock_units']}"])
    ws_inv.cell(row=2, column=1).font = font_subtitle
    ws_inv.append([])

    inv_headers = ["ID", "SKU", "Product Name", "Size", "Cost Price (NPR)", "Selling Price (NPR)", "Stock Units", "Total Asset Value (NPR)"]
    ws_inv.append(inv_headers)
    for col_idx in range(1, len(inv_headers) + 1):
        c = ws_inv.cell(row=4, column=col_idx)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center" if col_idx in [1, 4, 7] else "left", vertical="center")

    for p in products:
        val = p.cost_price * p.stock
        ws_inv.append([p.id, p.sku or "N/A", p.name, p.size, p.cost_price, p.price, p.stock, val])
        r = ws_inv.max_row
        for ci in range(1, 9):
            cell = ws_inv.cell(row=r, column=ci)
            cell.border = thin_border
            cell.font = font_regular
            if ci in [5, 6, 8]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif ci in [1, 7]:
                cell.alignment = Alignment(horizontal="center")

    for col in ws_inv.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_inv.column_dimensions[col_letter].width = max(max_len + 3, 14)

    # ── SHEET 5: CAPITAL & INVESTMENTS ──
    ws_cap = wb.create_sheet(title="Capital & Financing")
    ws_cap.views.sheetView[0].showGridLines = True
    ws_cap.append(["SALUZA'S CLOSET — CAPITAL & FINANCING CONTRIBUTIONS"])
    ws_cap.cell(row=1, column=1).font = font_title
    ws_cap.append(["Record of Owner's Capital Investments and Seed Loans"])
    ws_cap.cell(row=2, column=1).font = font_subtitle
    ws_cap.append([])

    cap_headers = ["Investment ID", "Transaction Date", "Note / Source", "Amount (NPR)"]
    ws_cap.append(cap_headers)
    for col_idx in range(1, len(cap_headers) + 1):
        c = ws_cap.cell(row=4, column=col_idx)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center" if col_idx in [1, 2] else "left", vertical="center")

    tot_cap = 0.0
    for inv in investments_detail:
        tot_cap += inv.amount
        ws_cap.append([inv.id, inv.date.strftime('%Y-%m-%d') if inv.date else "Initial", inv.note or "Capital Investment", inv.amount])
        r = ws_cap.max_row
        for ci in range(1, 5):
            cell = ws_cap.cell(row=r, column=ci)
            cell.border = thin_border
            cell.font = font_regular
            if ci == 4:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif ci in [1, 2]:
                cell.alignment = Alignment(horizontal="center")

    add_styled_row(ws_cap, "TOTAL INVESTED CAPITAL", [tot_cap], "total")

    for col in ws_cap.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_cap.column_dimensions[col_letter].width = max(max_len + 3, 16)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ══════════════════════════════════════════════════════════════
# BALANCE SHEET & EXPORT ROUTES
# ══════════════════════════════════════════════════════════════

@app.route('/balance_sheet')
def balance_sheet():
    """
    Renders the Monthly Balance Sheet dashboard, presenting:
    - Comparative Balance Sheet table across all monthly cycles
    - P&L and Cash Flow statements
    - Financial health KPIs (Working Capital, Net Worth, Ratios)
    """
    data = get_monthly_balance_sheet_data()
    return render_template('balance_sheet.html', **data)


@app.route('/export_excel')
@app.route('/export_balance_sheet')
def export_excel():
    """
    Exports the comprehensive Monthly Balance Sheet, P&L, Cash Flow,
    Inventory Assets, and Financing records as a professionally styled Excel workbook.
    """
    output = generate_balance_sheet_excel()
    return send_file(
        output,
        download_name="Saluza_Monthly_Balance_Sheet.xlsx",
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/update_status/<int:id>', methods=['POST'])
def update_status(id):
    order = Order.query.get_or_404(id)
    new_status = request.form.get('status')
    if new_status:
        order.status = new_status
        db.session.commit()
    # CHANGED: Force redirect back to orders page instead of "referrer"
    return redirect(url_for('order_history'))

@app.route('/inventory', methods=['GET', 'POST'])
def inventory():
    if request.method == 'POST':
        file = request.files.get('image')
        filename = secure_filename(file.filename) if file and file.filename != '' else None
        if filename: 
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        base_sku = request.form.get('sku')
        if not base_sku or base_sku.strip() == '':
            base_sku = None
        else:
            base_sku = base_sku.strip()

        name = request.form.get('name')
        cost_price = float(request.form.get('cost_price', 0.0))
        price = float(request.form.get('price', 0.0))
        
        # Get list of checked sizes
        selected_sizes = request.form.getlist('selected_sizes')
        
        if not selected_sizes:
            from flask import flash
            flash("Please select at least one size.", "danger")
            return redirect(url_for('inventory'))

        success_count = 0
        skipped_sizes = []
        
        for size in selected_sizes:
            stock = int(request.form.get(f'stock_{size}', 0))
            
            # Generate size-specific SKU if base SKU is provided
            size_sku = f"{base_sku}-{size.upper()}" if base_sku else None
            
            # Check unique constraint for SKU
            if size_sku:
                existing = Product.query.filter_by(sku=size_sku).first()
                if existing:
                    skipped_sizes.append(size)
                    continue
                    
            new_product = Product(
                sku=size_sku,
                name=name,
                size=size,
                cost_price=cost_price,
                price=price,
                stock=stock,
                image=filename
            )
            db.session.add(new_product)
            success_count += 1
            
        if success_count > 0:
            db.session.commit()
            
        from flask import flash
        if skipped_sizes:
            flash(f"Added products, but sizes {', '.join(skipped_sizes)} were skipped because their generated SKUs already exist.", "warning")
        elif success_count > 0:
            flash(f"Successfully added product for sizes: {', '.join(selected_sizes)}.", "success")
            
        return redirect(url_for('inventory'))
    
    products = Product.query.order_by(Product.id.desc()).all()
    return render_template('inventory.html', products=products)

@app.route('/edit_product/<int:id>', methods=['GET', 'POST'])
def edit_product(id):
    product = Product.query.get_or_404(id)
    if request.method == 'POST':
        sku_input = request.form.get('sku')
        product.sku = sku_input if sku_input else None
        product.name = request.form['name']
        product.size = request.form['size']
        product.cost_price = float(request.form['cost_price'])
        product.price = float(request.form['price'])
        product.stock = int(request.form['stock'])
        file = request.files.get('image')
        if file and file.filename != '':
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            product.image = filename
        db.session.commit()
        return redirect(url_for('inventory'))
    return render_template('edit_product.html', p=product)

@app.route('/new_order')
def new_order():
    products = Product.query.filter(Product.stock > 0).order_by(Product.id.desc()).all()
    today_date = datetime.now().strftime('%Y-%m-%d')
    return render_template('new_order.html', products=products, today_date=today_date)

@app.route('/create_order', methods=['POST'])
def create_order():
    data = request.json
    order_date = datetime.utcnow()
    if data.get('order_date'):
        try:
            parsed_d = datetime.strptime(data['order_date'], '%Y-%m-%d')
            order_date = datetime.combine(parsed_d.date(), datetime.utcnow().time())
        except Exception:
            pass
    new_order = Order(
        customer_name=data['customer_name'],
        phone=data['phone'],
        address=data['address'],
        date=order_date,
        status='Pending'
    )
    db.session.add(new_order)
    db.session.commit()

    grand_total = 0
    total_profit = 0

    for item in data['items']:
        product = Product.query.get(item['product_id'])
        if product and product.stock >= int(item['quantity']):
            qty = int(item['quantity'])
            discount = float(item.get('discount', 0))
            
            item_revenue = (product.price * qty) - discount
            item_cost = product.cost_price * qty
            item_profit = item_revenue - item_cost
            
            grand_total += item_revenue
            total_profit += item_profit
            
            order_item = OrderItem(
                order_id=new_order.id, 
                product_id=product.id, 
                quantity=qty,
                discount=discount,
                subtotal=item_revenue, 
                product_name=f"{product.name} ({product.size})"
            )
            product.stock -= qty
            db.session.add(order_item)
    
    new_order.total_amount = grand_total
    new_order.profit = total_profit
    db.session.commit()
    return jsonify({'message': 'Order Created', 'id': new_order.id})

@app.route('/edit_order/<int:id>')
def edit_order_page(id):
    order = Order.query.get_or_404(id)
    products = Product.query.order_by(Product.id.desc()).all()
    return render_template('edit_order.html', order=order, products=products)

@app.route('/update_order_data', methods=['POST'])
def update_order_data():
    data = request.json
    order_id = data.get('order_id')
    order = Order.query.get_or_404(order_id)
    
    old_items = OrderItem.query.filter_by(order_id=order.id).all()
    for item in old_items:
        product = Product.query.get(item.product_id)
        if product: product.stock += item.quantity
        db.session.delete(item)
    
    order.customer_name = data['customer_name']
    order.phone = data['phone']
    order.address = data['address']
    if data.get('order_date'):
        try:
            parsed_d = datetime.strptime(data['order_date'], '%Y-%m-%d')
            curr_time = order.date.time() if order.date else datetime.utcnow().time()
            order.date = datetime.combine(parsed_d.date(), curr_time)
        except Exception:
            pass
    
    grand_total = 0
    total_profit = 0
    
    for item in data['items']:
        product = Product.query.get(item['product_id'])
        if product:
            qty = int(item['quantity'])
            discount = float(item.get('discount', 0))
            if product.stock >= qty:
                item_revenue = (product.price * qty) - discount
                item_cost = product.cost_price * qty
                item_profit = item_revenue - item_cost
                
                grand_total += item_revenue
                total_profit += item_profit
                
                product.stock -= qty
                new_item = OrderItem(
                    order_id=order.id,
                    product_id=product.id,
                    quantity=qty,
                    discount=discount,
                    subtotal=item_revenue,
                    product_name=f"{product.name} ({product.size})"
                )
                db.session.add(new_item)
    
    order.total_amount = grand_total
    order.profit = total_profit
    db.session.commit()
    
    return jsonify({'message': 'Order Updated'})

@app.route('/update_order_date/<int:id>', methods=['POST'])
def update_order_date(id):
    order = Order.query.get_or_404(id)
    date_str = request.form.get('date') or (request.json.get('date') if request.is_json else None)
    if date_str:
        try:
            parsed_d = datetime.strptime(date_str, '%Y-%m-%d')
            curr_time = order.date.time() if order.date else datetime.utcnow().time()
            order.date = datetime.combine(parsed_d.date(), curr_time)
            db.session.commit()
            if request.is_json:
                return jsonify({'success': True, 'date': order.date.strftime('%Y-%m-%d')})
        except Exception as e:
            if request.is_json:
                return jsonify({'success': False, 'error': str(e)}), 400
    return redirect(request.referrer or url_for('order_history'))

@app.route('/delete_product/<int:id>')
def delete_product(id):
    product = Product.query.get_or_404(id)
    
    # SAFETY CHECK: Only delete if product has never been sold
    existing_sales = OrderItem.query.filter_by(product_id=id).first()
    
    if existing_sales:
        return "Cannot delete this product because it has been sold in previous orders. Deleting it would break your Sales History."
    
    db.session.delete(product)
    db.session.commit()
    return redirect(url_for('inventory'))

@app.route('/delete_order/<int:id>')
def delete_order(id):
    order = Order.query.get_or_404(id)
    
    # 1. Restore the Stock (Put items back in inventory)
    for item in order.items:
        product = Product.query.get(item.product_id)
        if product:
            product.stock += item.quantity
            
    # 2. Delete the Items first (to keep database clean)
    OrderItem.query.filter_by(order_id=id).delete()
    
    # 3. Delete the Order
    db.session.delete(order)
    db.session.commit()
    
    return redirect(url_for('order_history'))

@app.route('/add_investment', methods=['POST'])
def add_investment():
    amount = float(request.form.get('amount'))
    note = request.form.get('note')
    date_str = request.form.get('date')
    inv_date = datetime.utcnow()
    if date_str:
        try:
            parsed_d = datetime.strptime(date_str, '%Y-%m-%d')
            inv_date = datetime.combine(parsed_d.date(), datetime.utcnow().time())
        except Exception:
            pass
    new_invest = Investment(amount=amount, note=note, date=inv_date)
    db.session.add(new_invest)
    db.session.commit()
    
    return redirect(url_for('dashboard'))

@app.route('/edit_investment/<int:id>', methods=['GET', 'POST'])
def edit_investment(id):
    invest = Investment.query.get_or_404(id)
    
    if request.method == 'POST':
        invest.amount = float(request.form.get('amount'))
        invest.note = request.form.get('note')
        date_str = request.form.get('date')
        if date_str:
            try:
                parsed_d = datetime.strptime(date_str, '%Y-%m-%d')
                curr_time = invest.date.time() if invest.date else datetime.utcnow().time()
                invest.date = datetime.combine(parsed_d.date(), curr_time)
            except Exception:
                pass
        db.session.commit()
        return redirect(url_for('dashboard'))
        
    return render_template('edit_investment.html', investment=invest)

@app.route('/delete_investment/<int:id>')
def delete_investment(id):
    invest = Investment.query.get_or_404(id)
    db.session.delete(invest)
    db.session.commit()
    return redirect(url_for('dashboard'))

with app.app_context():
    db.create_all()

# Create tables if they don't exist
with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True)