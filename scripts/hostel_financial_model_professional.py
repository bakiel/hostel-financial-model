', end_color='366092', fill_type='solid')
            ws.cell(row=row, column=col).font = Font(color='FFFFFF', bold=True)
        
        row += 1
        for room_type, config in self.room_configuration.items():
            ws.cell(row=row, column=1, value=room_type.replace('_', ' ').title())
            ws.cell(row=row, column=2, value=config['beds'])
            ws.cell(row=row, column=3, value=f"${config['rate']}")
            ws.cell(row=row, column=4, value=config['category'].title())
            ws.cell(row=row, column=5, value=', '.join(config['amenities']))
            row += 1
        
        # Total summary
        total_beds = sum(config['beds'] for config in self.room_configuration.values())
        avg_rate = sum(config['beds'] * config['rate'] for config in self.room_configuration.values()) / total_beds
        
        row += 1
        ws.cell(row=row, column=1, value='TOTAL')
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=2, value=total_beds)
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=3, value=f"${avg_rate:.2f} (weighted avg)")
        ws.cell(row=row, column=3).font = Font(bold=True)
    
    def _add_financial_assumptions(self, ws):
        """Add financial assumptions"""
        row = 15
        ws.cell(row=row, column=1, value='Financial Assumptions')
        ws.cell(row=row, column=1).font = Font(size=12, bold=True)
        
        assumptions = [
            ('Growth & Inflation', [
                ('Annual Revenue Growth', '3.0%'),
                ('Expense Inflation Rate', '2.5%'),
                ('Wage Growth Rate', '3.0%'),
                ('Rate Growth Annual', '2.0%')
            ]),
            ('Operating Assumptions', [
                ('Base Occupancy - Low Season', '60%'),
                ('Base Occupancy - Mid Season', '70%'),
                ('Base Occupancy - High Season', '80%'),
                ('Base Occupancy - Peak Season', '90%'),
                ('OTA Commission Rate', '15%'),
                ('Direct Booking Target', '40%')
            ]),
            ('Financial Parameters', [
                ('Tax Rate', '25%'),
                ('Discount Rate (WACC)', '10%'),
                ('Terminal Growth Rate', '2%'),
                ('Depreciation Period', '10 years'),
                ('Loan Interest Rate', '6%'),
                ('Loan Term', '5 years')
            ])
        ]
        
        row += 2
        for section, items in assumptions:
            ws.cell(row=row, column=1, value=section)
            ws.cell(row=row, column=1).font = Font(bold=True, underline='single')
            row += 1
            
            for param, value in items:
                ws.cell(row=row, column=1, value=param)
                ws.cell(row=row, column=3, value=value)
                ws.cell(row=row, column=3).font = Font(bold=True)
                row += 1
            
            row += 1
    
    def _add_market_assumptions(self, ws):
        """Add market assumptions"""
        row = 40
        ws.cell(row=row, column=1, value='Market Assumptions')
        ws.cell(row=row, column=1).font = Font(size=12, bold=True)
        
        market_assumptions = [
            ('Total Market Size', '500,000 room nights/year'),
            ('Market Growth Rate', '5% annually'),
            ('Current Market Share', '2.0%'),
            ('Target Market Share (Y5)', '4.0%'),
            ('Primary Customer Segments', 'Backpackers (40%), Digital Nomads (30%), Weekend Travelers (30%)'),
            ('Average Length of Stay', '2.8 nights'),
            ('Booking Window', '21 days advance'),
            ('Seasonality Impact', '±30% from average')
        ]
        
        row += 2
        for assumption, value in market_assumptions:
            ws.cell(row=row, column=1, value=assumption)
            ws.cell(row=row, column=3, value=value)
            ws.cell(row=row, column=3).font = Font(bold=True)
            row += 1
    
    def _create_monthly_details(self, writer):
        """Create detailed monthly projections for 60 months"""
        ws = writer.book.create_sheet('Monthly Details')
        
        # Title
        ws['A1'] = '60-MONTH DETAILED PROJECTIONS'
        ws['A1'].font = Font(size=16, bold=True)
        
        # Generate 60 months of data
        monthly_data = self._generate_60_month_projections()
        
        # Write data
        row = 4
        for r in dataframe_to_rows(monthly_data, index=False, header=True):
            for col, value in enumerate(r, 1):
                ws.cell(row=row, column=col, value=value)
                if row == 4:  # Header row
                    ws.cell(row=row, column=col).font = Font(bold=True)
                    ws.cell(row=row, column=col).fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
                    ws.cell(row=row, column=col).font = Font(color='FFFFFF', bold=True)
            row += 1
        
        # Format numbers
        for row in range(5, ws.max_row + 1):
            for col in range(4, ws.max_column + 1):
                if ws.cell(row=row, column=col).value and isinstance(ws.cell(row=row, column=col).value, (int, float)):
                    if col in [10, 11]:  # Percentage columns
                        ws.cell(row=row, column=col).number_format = '0.0%'
                    else:
                        ws.cell(row=row, column=col).number_format = '"$"#,##0'
        
        # Add conditional formatting for Net Income
        net_income_col = 9  # Assuming Net Income is column 9
        ws.conditional_formatting.add(
            f'{get_column_letter(net_income_col)}5:{get_column_letter(net_income_col)}{ws.max_row}',
            ColorScaleRule(
                start_type='min',
                start_color='FF6B6B',  # Red
                mid_type='percentile',
                mid_value=50,
                mid_color='FFE66D',     # Yellow
                end_type='max',
                end_color='51CF66'      # Green
            )
        )
    
    def _generate_60_month_projections(self):
        """Generate 60 months of detailed projections"""
        projections = []
        
        for month_num in range(60):
            date = self.start_date + timedelta(days=30 * month_num)
            year_offset = month_num // 12
            month = date.month
            
            # Revenue calculation
            revenue_data = self._calculate_detailed_revenue(date.year, month, year_offset)
            
            # Expense calculation
            expense_data = self._calculate_detailed_expenses(
                revenue_data['total_revenue'],
                revenue_data['occupancy'],
                date.year,
                month,
                year_offset
            )
            
            # Financial metrics
            ebitda = revenue_data['total_revenue'] - expense_data['total_expenses']
            depreciation = self.investment_params['initial_investment'] / 120  # Monthly depreciation
            interest = max(0, (self.investment_params['initial_investment'] * 0.5 - 10000 * month_num) * 0.06 / 12)
            tax = max(0, (ebitda - depreciation - interest) * 0.25)
            net_income = ebitda - depreciation - interest - tax
            
            projection = {
                'Month_Num': month_num + 1,
                'Date': date.strftime('%Y-%m'),
                'Year': date.year,
                'Month': date.strftime('%B'),
                'Room_Revenue': revenue_data['room_revenue'],
                'Other_Revenue': revenue_data['total_revenue'] - revenue_data['room_revenue'],
                'Total_Revenue': revenue_data['total_revenue'],
                'Total_Expenses': expense_data['total_expenses'],
                'EBITDA': ebitda,
                'Net_Income': net_income,
                'Occupancy': revenue_data['occupancy'],
                'Cumulative_Revenue': 0,  # Will calculate after
                'Cumulative_NI': 0  # Will calculate after
            }
            
            projections.append(projection)
        
        # Calculate cumulative values
        df = pd.DataFrame(projections)
        df['Cumulative_Revenue'] = df['Total_Revenue'].cumsum()
        df['Cumulative_NI'] = df['Net_Income'].cumsum()
        
        return df
    
    def _create_charts_sheet(self, writer):
        """Create sheet with various charts and visualizations"""
        ws = writer.book.create_sheet('Charts')
        
        # Title
        ws['A1'] = 'FINANCIAL VISUALIZATIONS'
        ws['A1'].font = Font(size=16, bold=True)
        
        # Revenue breakdown pie chart
        self._add_revenue_breakdown_chart(ws)
        
        # Expense breakdown pie chart
        self._add_expense_breakdown_chart(ws)
        
        # Trend charts
        self._add_trend_charts(ws)
        
        # Occupancy heat map data
        self._add_occupancy_heatmap(ws)
    
    def _add_revenue_breakdown_chart(self, ws):
        """Add revenue breakdown pie chart"""
        # Data for chart
        row = 4
        ws.cell(row=row, column=1, value='Revenue Breakdown (Year 1)')
        ws.cell(row=row, column=1).font = Font(size=12, bold=True)
        
        revenue_categories = [
            ('Room Revenue - Dorms', 180000),
            ('Room Revenue - Private', 120000),
            ('F&B Revenue', 45000),
            ('Activities & Tours', 30000),
            ('Other Services', 15000)
        ]
        
        row += 2
        ws.cell(row=row, column=1, value='Category')
        ws.cell(row=row, column=2, value='Amount')
        
        row += 1
        for category, amount in revenue_categories:
            ws.cell(row=row, column=1, value=category)
            ws.cell(row=row, column=2, value=amount)
            row += 1
        
        # Create pie chart
        pie = PieChart()
        pie.title = "Revenue Sources - Year 1"
        labels = Reference(ws, min_col=1, min_row=7, max_row=11)
        data = Reference(ws, min_col=2, min_row=6, max_row=11)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        pie.height = 10
        pie.width = 15
        
        ws.add_chart(pie, "D4")
    
    def _add_expense_breakdown_chart(self, ws):
        """Add expense breakdown pie chart"""
        # Data for chart
        row = 20
        ws.cell(row=row, column=1, value='Expense Breakdown (Year 1)')
        ws.cell(row=row, column=1).font = Font(size=12, bold=True)
        
        expense_categories = [
            ('Staff Costs', 150000),
            ('Utilities', 24000),
            ('Marketing', 36000),
            ('Supplies', 18000),
            ('Maintenance', 15000),
            ('Insurance', 9600),
            ('Other', 12000)
        ]
        
        row += 2
        ws.cell(row=row, column=1, value='Category')
        ws.cell(row=row, column=2, value='Amount')
        
        row += 1
        for category, amount in expense_categories:
            ws.cell(row=row, column=1, value=category)
            ws.cell(row=row, column=2, value=amount)
            row += 1
        
        # Create pie chart
        pie2 = PieChart()
        pie2.title = "Operating Expenses - Year 1"
        labels2 = Reference(ws, min_col=1, min_row=23, max_row=29)
        data2 = Reference(ws, min_col=2, min_row=22, max_row=29)
        pie2.add_data(data2, titles_from_data=True)
        pie2.set_categories(labels2)
        pie2.height = 10
        pie2.width = 15
        
        ws.add_chart(pie2, "D20")
    
    def _add_trend_charts(self, ws):
        """Add trend analysis charts"""
        # Monthly revenue trend
        row = 4
        col = 13
        ws.cell(row=row, column=col, value='Monthly Trends (First 24 months)')
        ws.cell(row=row, column=col).font = Font(size=12, bold=True)
        
        # Sample data for trends
        row += 2
        ws.cell(row=row, column=col, value='Month')
        ws.cell(row=row, column=col+1, value='Revenue')
        ws.cell(row=row, column=col+2, value='Expenses')
        ws.cell(row=row, column=col+3, value='Occupancy')
        
        # Generate trend data
        for i in range(24):
            row += 1
            base_revenue = 30000 + i * 500
            seasonality = 1 + 0.2 * np.sin(i * np.pi / 6)
            
            ws.cell(row=row, column=col, value=i+1)
            ws.cell(row=row, column=col+1, value=base_revenue * seasonality)
            ws.cell(row=row, column=col+2, value=base_revenue * 0.7)
            ws.cell(row=row, column=col+3, value=0.65 + 0.15 * seasonality)
        
        # Create combination chart
        chart3 = LineChart()
        chart3.title = "Revenue & Expense Trends"
        chart3.y_axis.title = "Amount ($)"
        chart3.x_axis.title = "Month"
        
        # Add data series
        values1 = Reference(ws, min_col=col+1, min_row=6, max_row=30)
        values2 = Reference(ws, min_col=col+2, min_row=6, max_row=30)
        categories = Reference(ws, min_col=col, min_row=7, max_row=30)
        
        chart3.add_data(values1, titles_from_data=True)
        chart3.add_data(values2, titles_from_data=True)
        chart3.set_categories(categories)
        
        chart3.height = 12
        chart3.width = 20
        
        ws.add_chart(chart3, "M7")
    
    def _add_occupancy_heatmap(self, ws):
        """Add occupancy heatmap data"""
        row = 35
        ws.cell(row=row, column=1, value='Occupancy Heat Map by Month and Room Type')
        ws.cell(row=row, column=1).font = Font(size=12, bold=True)
        
        # Create heatmap data
        row += 2
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        room_types = ['4-Bed Dorm', '6-Bed Dorm', 'Private Single', 'Private Double']
        
        # Headers
        ws.cell(row=row, column=1, value='Room Type')
        for col, month in enumerate(months, 2):
            ws.cell(row=row, column=col, value=month)
            ws.cell(row=row, column=col).font = Font(bold=True)
        
        # Data with conditional formatting
        row += 1
        for room_type in room_types:
            ws.cell(row=row, column=1, value=room_type)
            ws.cell(row=row, column=1).font = Font(bold=True)
            
            for col, month_idx in enumerate(range(12), 2):
                # Generate occupancy based on seasonality
                season_data = self.seasonality_patterns[month_idx + 1]
                base_occ = self.base_occupancy[season_data['season']]
                occupancy = base_occ + season_data['adjustment'] + np.random.uniform(-0.05, 0.05)
                
                ws.cell(row=row, column=col, value=occupancy)
                ws.cell(row=row, column=col).number_format = '0%'
                
                # Color based on occupancy
                if occupancy >= 0.85:
                    color = '00B050'  # Dark green
                elif occupancy >= 0.75:
                    color = '92D050'  # Light green
                elif occupancy >= 0.65:
                    color = 'FFFF00'  # Yellow
                elif occupancy >= 0.55:
                    color = 'FFC000'  # Orange
                else:
                    color = 'FF0000'  # Red
                
                ws.cell(row=row, column=col).fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
            
            row += 1
    
    def _apply_professional_formatting(self, writer):
        """Apply professional formatting to all sheets"""
        workbook = writer.book
        
        # Define named styles
        header_style = NamedStyle(name='header_style')
        header_style.font = Font(bold=True, color='FFFFFF')
        header_style.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_style.alignment = Alignment(horizontal='center', vertical='center')
        
        currency_style = NamedStyle(name='currency_style')
        currency_style.number_format = '"$"#,##0'
        
        percent_style = NamedStyle(name='percent_style')
        percent_style.number_format = '0.0%'
        
        # Apply to all sheets
        for sheet in workbook.worksheets:
            # Set default column width
            for column in sheet.columns:
                max_length = 0
                column_letter = None
                
                for cell in column:
                    if hasattr(cell, 'column_letter'):
                        column_letter = cell.column_letter
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                
                if column_letter:
                    adjusted_width = min(max_length + 2, 30)
                    sheet.column_dimensions[column_letter].width = adjusted_width
            
            # Freeze panes for data sheets
            if sheet.title not in ['Cover', 'Contents']:
                sheet.freeze_panes = 'A5'
            
            # Set print settings
            sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.fitToHeight = False
            sheet.print_options.horizontalCentered = True
            sheet.print_options.verticalCentered = True
            
            # Add headers and footers
            sheet.oddHeader.center.text = f"&B{self.hostel_name} - Financial Model"
            sheet.oddHeader.center.font = "Arial,Bold"
            sheet.oddHeader.center.size = 12
            
            sheet.oddFooter.left.text = "&D &T"
            sheet.oddFooter.center.text = "Page &P of &N"
            sheet.oddFooter.right.text = "Confidential"


# Main execution
if __name__ == "__main__":
    print("Creating Professional Hostel Financial Model...")
    print("This will generate a comprehensive 100KB+ Excel file with advanced analytics...")
    
    model = ProfessionalHostelFinancialModel()
    filename = model.create_comprehensive_model()
    
    print(f"\nModel generation complete!")
    print(f"File created: {filename}")
    print("\nThe model includes:")
    print("- 20+ worksheets with detailed analysis")
    print("- 5-year financial projections")
    print("- Scenario and sensitivity analysis")
    print("- Monte Carlo risk simulation")
    print("- Market and competitor analysis")
    print("- Investment metrics and KPIs")
    print("- Professional charts and visualizations")
