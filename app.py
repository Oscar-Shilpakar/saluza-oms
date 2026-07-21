import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta

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

    # For display
    net_profit = gross_profit # As you requested previously (Sales Profit only)
    
    recent_orders = Order.query.order_by(Order.date.desc()).limit(5).all()
    investments = Investment.query.order_by(Investment.date.desc()).all()

    # --- 4. BUSINESS INSIGHTS ---
    total_orders_count = Order.query.count()
    avg_order_value = (total_sales / total_orders_count) if total_orders_count else 0
    profit_margin_pct = (gross_profit / total_sales * 100) if total_sales else 0
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

# --- UPDATED EXPENSES ROUTE ---
@app.route('/expenses', methods=['GET', 'POST'])
def expenses():
    # Handle Manual Expense Addition
    if request.method == 'POST':
        new_expense = Expense(
            description=request.form['description'],
            category=request.form['category'],
            amount=float(request.form['amount'])
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
    
    return render_template('expenses.html', 
                           expenses=all_expenses, 
                           product_expense=total_product_expense)

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

# --- REPLACEMENT FOR THE EXPORT ROUTE ---
@app.route('/export_excel')
def export_excel():
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # --- SHEET 1: ORDER SUMMARY (The "Headers") ---
    orders = Order.query.all()
    order_data = []
    for o in orders:
        order_data.append({
            "Order ID": o.id,
            "Date": o.date.strftime('%Y-%m-%d %H:%M'),
            "Customer Name": o.customer_name,
            "Phone": o.phone,
            "Address": o.address,
            "Total Amount": o.total_amount,
            "Calculated Profit": o.profit,
            "Status": o.status
        })
    pd.DataFrame(order_data).to_excel(writer, index=False, sheet_name='All Orders')

    # --- SHEET 2: SOLD ITEMS (The "Details") ---
    items = OrderItem.query.all()
    item_data = []
    for i in items:
        # Get parent order info for reference
        parent_order = Order.query.get(i.order_id)
        order_date = parent_order.date.strftime('%Y-%m-%d') if parent_order else "Unknown"
        
        item_data.append({
            "Order ID": i.order_id,
            "Order Date": order_date,
            "Product Name": i.product_name,
            "Quantity": i.quantity,
            "Discount Given": i.discount,
            "Subtotal (Revenue)": i.subtotal
        })
    pd.DataFrame(item_data).to_excel(writer, index=False, sheet_name='Sold Items Detail')

    # --- SHEET 3: CURRENT INVENTORY (Your Stock) ---
    products = Product.query.all()
    prod_data = []
    total_inventory_value = 0
    for p in products:
        stock_value = p.stock * p.cost_price
        total_inventory_value += stock_value
        prod_data.append({
            "ID": p.id,
            "SKU": p.sku,
            "Name": p.name,
            "Size": p.size,
            "Cost Price": p.cost_price,
            "Selling Price": p.price,
            "Current Stock": p.stock,
            "Stock Value (Asset)": stock_value,
            "Image File": p.image
        })
    pd.DataFrame(prod_data).to_excel(writer, index=False, sheet_name='Current Inventory')

    # --- SHEET 4: MANUAL EXPENSES (Rent, Ads) ---
    expenses = Expense.query.all()
    exp_data = []
    total_manual_expenses = 0
    for e in expenses:
        total_manual_expenses += e.amount
        exp_data.append({
            "ID": e.id,
            "Date": e.date.strftime('%Y-%m-%d'),
            "Category": e.category,
            "Description": e.description,
            "Amount": e.amount
        })
    pd.DataFrame(exp_data).to_excel(writer, index=False, sheet_name='Other Expenses')

    # --- SHEET 5: FINANCIAL SNAPSHOT (The Dashboard Numbers) ---
    # Calculate Sold Goods Cost
    sold_items = OrderItem.query.all()
    sold_cost = 0
    for i in sold_items:
        product = Product.query.get(i.product_id)
        if product:
            sold_cost += (i.quantity * product.cost_price)
            
    total_sales = db.session.query(db.func.sum(Order.total_amount)).scalar() or 0
    
    # Summary Table
    summary_data = [
        {"Metric": "Total Sales Revenue", "Value": total_sales},
        {"Metric": "Total Manual Expenses (Rent/Ads)", "Value": total_manual_expenses},
        {"Metric": "Value of Unsold Inventory", "Value": total_inventory_value},
        {"Metric": "Cost of Sold Goods", "Value": sold_cost},
        {"Metric": "Total Product Investment (Sold + Unsold)", "Value": total_inventory_value + sold_cost},
        {"Metric": "NET PROFIT (Cash Basis)", "Value": total_sales - (total_manual_expenses + total_inventory_value + sold_cost)}
    ]
    pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='Financial Summary')

    writer.close()
    output.seek(0)
    
    return send_file(output, download_name="Saluza_Full_Data_Export.xlsx", as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

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
    
    products = Product.query.all()
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
    products = Product.query.filter(Product.stock > 0).all()
    return render_template('new_order.html', products=products)

@app.route('/create_order', methods=['POST'])
def create_order():
    data = request.json
    new_order = Order(
        customer_name=data['customer_name'],
        phone=data['phone'],
        address=data['address'],
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
    products = Product.query.all()
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
    
    new_invest = Investment(amount=amount, note=note)
    db.session.add(new_invest)
    db.session.commit()
    
    return redirect(url_for('dashboard'))

@app.route('/edit_investment/<int:id>', methods=['GET', 'POST'])
def edit_investment(id):
    invest = Investment.query.get_or_404(id)
    
    if request.method == 'POST':
        invest.amount = float(request.form.get('amount'))
        invest.note = request.form.get('note')
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