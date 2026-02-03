import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
from datetime import datetime

app = Flask(__name__)

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
    products = Product.query.all()
    for p in products:
        current_inventory_cost += (p.stock * p.cost_price)
        
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
    
    # Queries for lists
    recent_orders = Order.query.order_by(Order.date.desc()).limit(5).all()
    low_stock = Product.query.filter(Product.stock < 2, Product.stock > 0).all()
    out_of_stock = Product.query.filter(Product.stock == 0).all()
    investments = Investment.query.order_by(Investment.date.desc()).all()
    
    return render_template('dashboard.html', 
                           sales=total_sales, 
                           gross_profit=gross_profit, 
                           expenses=manual_expenses, 
                           net_profit=net_profit, 
                           cash_in_hand=cash_in_hand, # <--- NEW DATA
                           total_investment=total_investment, # <--- NEW DATA
                           orders=recent_orders,
                           investments=investments, 
                           low_stock=low_stock,
                           out_of_stock=out_of_stock)

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
        filename = secure_filename(file.filename) if file else None
        if filename: file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        sku_input = request.form.get('sku')
        if not sku_input: sku_input = None 

        new_product = Product(
            sku=sku_input,
            name=request.form['name'],
            size=request.form['size'],
            cost_price=float(request.form['cost_price']),
            price=float(request.form['price']),
            stock=int(request.form['stock']),
            image=filename
        )
        db.session.add(new_product)
        db.session.commit()
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