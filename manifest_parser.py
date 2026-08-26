import streamlit as st
import pandas as pd
import re
from io import BytesIO
import xlsxwriter
import csv
import os
from datetime import datetime
from PIL import Image
import plotly.graph_objects as go
import plotly.express as px

def parse_manifest_csv(file_content, filename):
    """Parse the manifest CSV file content and extract structured data"""
    
    # Read the CSV file
    content_str = file_content.read().decode('utf-8')
    lines = content_str.splitlines()
    
    # Parse CSV with proper handling of quoted fields
    rows = []
    reader = csv.reader(lines, quotechar='"', delimiter=',', quoting=csv.QUOTE_MINIMAL)
    for row in reader:
        rows.append(row)
    
    # Extract trip type from first row, first cell
    trip_type = "Morning"  # Default
    if rows and len(rows) > 0:
        first_cell = rows[0][0] if len(rows[0]) > 0 else ""
        if first_cell:
            # Look for Morning or Afternoon in the first cell
            if "Morning" in first_cell:
                trip_type = "Morning"
            elif "Afternoon" in first_cell:
                trip_type = "Afternoon"
            elif "morning" in first_cell.lower():
                trip_type = "Morning"
            elif "afternoon" in first_cell.lower():
                trip_type = "Afternoon"
    
    bookings = []
    
    # Find the header row that contains "Booking ID"
    header_row_index = None
    booking_id_col = None
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            if "Booking ID" in str(cell):
                header_row_index = i
                booking_id_col = j
                break
        if header_row_index is not None:
            break
    
    if header_row_index is None:
        st.error("Could not find header row with 'Booking ID' in CSV file")
        return None, trip_type
    
    # Find the summary row (contains "bookings")
    summary_row_index = None
    for i, row in enumerate(rows):
        for cell in row:
            if cell and "bookings" in str(cell).lower():
                summary_row_index = i
                break
        if summary_row_index is not None:
            break
    
    if summary_row_index is None:
        st.error("Could not find summary row in CSV file")
        return None, trip_type
    
    # Extract total bookings from summary row
    total_bookings = 0
    for cell in rows[summary_row_index]:
        if cell and "bookings" in cell.lower():
            match = re.search(r'(\d+)\s*bookings', cell.lower())
            if match:
                total_bookings = int(match.group(1))
                break
    
    # Define column indices based on header row
    header_cells = rows[header_row_index]
    column_indices = {}
    for idx, header in enumerate(header_cells):
        header_clean = header.strip().lower()
        if 'booking' in header_clean and 'id' in header_clean:
            column_indices['booking_id'] = idx
        elif 'contact' in header_clean:
            column_indices['contact'] = idx
        elif 'phone' in header_clean and 'number' not in header_clean:
            column_indices['phone'] = idx
        elif 'email' in header_clean:
            column_indices['email'] = idx
        elif 'customers' in header_clean:
            column_indices['customers'] = idx
        elif 'pax' in header_clean:
            column_indices['pax'] = idx
        elif 'total' in header_clean and 'pay' not in header_clean:
            column_indices['total'] = idx
        elif 'due' in header_clean:
            column_indices['due'] = idx
        elif 'created' in header_clean:
            column_indices['created_by'] = idx
        elif 'notes' in header_clean:
            column_indices['notes'] = idx
        elif 'transport' in header_clean:
            column_indices['transport'] = idx
        elif 'comments' in header_clean:
            column_indices['comments'] = idx
    
    # Now parse bookings starting from after the summary row
    current_row = summary_row_index + 1
    booking_count = 0
    
    while booking_count < total_bookings and current_row < len(rows):
        # Skip empty rows
        if not any(cell for cell in rows[current_row] if cell and str(cell).strip()):
            current_row += 1
            continue
        
        # Check if this row starts a booking (starts with #)
        booking_id_idx = column_indices.get('booking_id', 0)
        booking_id = rows[current_row][booking_id_idx] if len(rows[current_row]) > booking_id_idx else ""
        
        if booking_id and booking_id.startswith('#'):
            # This is a booking header row
            booking_row = rows[current_row]
            
            # Move to next row for passenger details
            passenger_row_idx = current_row + 1
            if passenger_row_idx >= len(rows):
                break
            
            # Passenger details are in the next row (first cell)
            passenger_details_row = rows[passenger_row_idx]
            
            # Parse booking header
            booking = parse_booking_header(booking_row, column_indices)
            
            # Get the number of passengers from the booking
            try:
                num_passengers = int(booking.get('pax', '0'))
            except (ValueError, TypeError):
                num_passengers = 0
            
            # Parse passenger details
            passengers = []
            if passenger_details_row and passenger_details_row[0]:
                # The passenger details are in the first cell of this row
                passenger_details_text = passenger_details_row[0]
                passengers = parse_passenger_details_flexible(passenger_details_text, num_passengers)
            
            # If we couldn't parse enough passengers, create empty ones
            if len(passengers) < num_passengers:
                for i in range(num_passengers - len(passengers)):
                    passengers.append(create_empty_passenger())
            
            booking['passengers'] = passengers
            bookings.append(booking)
            booking_count += 1
            
            # Skip the passenger details row(s) we just processed
            current_row = passenger_row_idx + 1
        else:
            current_row += 1
    
    return bookings, trip_type

def parse_booking_header(booking_row, column_indices):
    """Parse the booking header row"""
    booking = {
        'booking_id': "",
        'contact_name': "",
        'phone': "",
        'email': "",
        'customers': "",
        'pax': "",
        'total': "",
        'due': "",
        'source': "",  # This comes from "Created by" column
        'source_code': "",  # This comes from "Notes" column
        'passengers': [],
        'comments': "",
        'transport': "No",
        'transport_details': "",
        'pickup_location': "",  # New field for transport location
        'pickup_time': ""       # New field for pickup time
    }
    
    # Extract values using column indices
    try:
        booking['booking_id'] = booking_row[column_indices.get('booking_id', 0)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['contact_name'] = booking_row[column_indices.get('contact', 1)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['phone'] = booking_row[column_indices.get('phone', 2)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['email'] = booking_row[column_indices.get('email', 3)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['customers'] = booking_row[column_indices.get('customers', 4)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['pax'] = booking_row[column_indices.get('pax', 5)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['total'] = booking_row[column_indices.get('total', 10)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['due'] = booking_row[column_indices.get('due', 11)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['source'] = booking_row[column_indices.get('created_by', 12)].strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        notes = booking_row[column_indices.get('notes', 13)].strip()
        if "znb" in notes.lower():
            source_code = notes.lower().replace("znb", "").strip().upper()
            booking['source_code'] = source_code
        elif notes:
            booking['source_code'] = notes.strip()
    except (IndexError, AttributeError):
        pass
    
    try:
        transport = booking_row[column_indices.get('transport', 8)].strip()
        booking['transport'] = "Yes" if transport.lower() == "yes" else "No"
    except (IndexError, AttributeError):
        pass
    
    try:
        booking['comments'] = booking_row[column_indices.get('comments', 7)].strip()
    except (IndexError, AttributeError):
        pass
    
    return booking

def parse_passenger_details_flexible(passenger_text, expected_passenger_count):
    """Parse passenger details flexibly - handle any field order and missing fields"""
    passengers = []
    
    if not passenger_text:
        # If no passenger text, create empty passengers based on expected count
        for i in range(expected_passenger_count):
            passengers.append(create_empty_passenger())
        return passengers
    
    # Extract transport details from passenger text if available
    pickup_location = ""
    pickup_time = ""
    
    # Look for pickup location pattern
    pickup_match = re.search(r'Pickup location:\s*([^\n]+?)\s*(?:Pickup time:|$)', passenger_text)
    if pickup_match:
        pickup_location = pickup_match.group(1).strip()
    
    # Look for pickup time pattern
    time_match = re.search(r'Pickup time:\s*([^\n]+?)\s*(?:Edit|$)', passenger_text)
    if time_match:
        pickup_time = time_match.group(1).strip()
    
    # Store transport details in a variable to add to booking later
    transport_details = {
        'pickup_location': pickup_location,
        'pickup_time': pickup_time
    }
    
    # First, let's identify passenger blocks more flexibly
    # Look for patterns that indicate a new passenger
    passenger_patterns = [
        r'Adult\s*(?:Full Name:|Name:|Gender:|Date of Birth:|Phone number:|Passport or ID number:|Country List:)',
        r'Child\s*(?:Full Name:|Name:|Gender:|Date of Birth:|Phone number:|Passport or ID number:|Country List:)',
        r'Infant\s*(?:Full Name:|Name:|Gender:|Date of Birth:|Phone number:|Passport or ID number:|Country List:)'
    ]
    
    # Find all starting positions of new passengers
    start_positions = []
    for pattern in passenger_patterns:
        for match in re.finditer(pattern, passenger_text):
            start_positions.append(match.start())
    
    # Sort start positions
    start_positions.sort()
    
    # If no start positions found, try to extract at least one passenger
    if not start_positions:
        passenger = parse_single_passenger(passenger_text)
        if passenger['full_name']:  # Only add if we found a name
            passengers.append(passenger)
    else:
        # Extract each passenger block
        for i, start_pos in enumerate(start_positions):
            if i < len(start_positions) - 1:
                end_pos = start_positions[i + 1]
                passenger_block = passenger_text[start_pos:end_pos]
            else:
                passenger_block = passenger_text[start_pos:]
            
            passenger = parse_single_passenger(passenger_block)
            passengers.append(passenger)
    
    # If we still have fewer passengers than expected, add empty ones
    while len(passengers) < expected_passenger_count:
        passengers.append(create_empty_passenger())
    
    # Add transport details to each passenger for now (will be consolidated later)
    for passenger in passengers:
        passenger['pickup_location'] = pickup_location
        passenger['pickup_time'] = pickup_time
    
    return passengers[:expected_passenger_count]  # Don't exceed expected count

def parse_single_passenger(passenger_block):
    """Parse a single passenger block with any field order"""
    passenger = {
        'full_name': "",
        'gender': "",
        'phone': "",
        'nationality': "",
        'dob': "",
        'passport': "",
        'type': "Adult"  # Default
    }
    
    # Determine passenger type
    if 'Child' in passenger_block[:10]:
        passenger['type'] = "Child"
    elif 'Infant' in passenger_block[:10]:
        passenger['type'] = "Infant"
    
    # Parse fields in any order using regex patterns
    field_patterns = {
        'full_name': [
            r'Full Name:\s*([^E]+?)\s*Edit',
            r'Full Name:\s*([^\n]+?)\s*(?:Edit|$)',
            r'Name:\s*([^E]+?)\s*Edit',
            r'([^E]+?)\s*Edit Date of Birth:'
        ],
        'gender': [
            r'Gender:\s*([^E]+?)\s*Edit',
            r'Gender:\s*([^\n]+?)\s*(?:Edit|$)'
        ],
        'phone': [
            r'Phone number:\s*([^E]+?)\s*Edit',
            r'Phone number:\s*([^\n]+?)\s*(?:Edit|$)',
            r'Phone:\s*([^E]+?)\s*Edit'
        ],
        'nationality': [
            r'Country List:\s*([^E]+?)\s*Edit',
            r'Country List:\s*([^\n]+?)\s*(?:Edit|$)',
            r'Country:\s*([^E]+?)\s*Edit'
        ],
        'dob': [
            r'Date of Birth:\s*([^E]+?)\s*Edit',
            r'Date of Birth:\s*([^\n]+?)\s*(?:Edit|$)',
            r'Birth:\s*([^E]+?)\s*Edit'
        ],
        'passport': [
            r'Passport or ID number:\s*([^E]+?)\s*Edit',
            r'Passport or ID number:\s*([^\n]+?)\s*(?:Edit|$)',
            r'Passport:\s*([^E]+?)\s*Edit',
            r'ID number:\s*([^E]+?)\s*Edit'
        ]
    }
    
    # Try each pattern for each field
    for field, patterns in field_patterns.items():
        for pattern in patterns:
            match = re.search(pattern, passenger_block, re.IGNORECASE | re.DOTALL)
            if match:
                value = match.group(1).strip()
                # Clean up the value
                if value.endswith('Edit'):
                    value = value[:-4].strip()
                passenger[field] = value
                break
    
    # Clean up gender field
    if passenger['gender']:
        gender = passenger['gender'].lower()
        if gender == 'male':
            passenger['gender'] = 'M'
        elif gender == 'female':
            passenger['gender'] = 'F'
        elif len(gender) > 1:
            passenger['gender'] = gender[0].upper()
    
    return passenger

def create_empty_passenger():
    """Create an empty passenger record"""
    return {
        'full_name': "",
        'gender': "",
        'phone': "",
        'nationality': "",
        'dob': "",
        'passport': "",
        'type': "Adult",
        'pickup_location': "",
        'pickup_time': ""
    }

def create_output_file(bookings, tour_date, trip_type):
    """Create output Excel file in the desired format with color-coded booking groups"""
    
    output = BytesIO()
    
    # Create workbook
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Passenger List')
    
    # Define color palette for booking groups (light alternating colors)
    color_palette = [
        "#C0D5E8",  # Light Blue
        '#FFFFFF',  # White
    ]
    
    # Define special formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#D9EAD3'  # Light green header
    })
    
    # Create special formats for infant, child, and transport
    infant_cells_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter',
        'bg_color': '#CCE5FF'  # Light blue for infant cells
    })
    
    child_cells_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter',
        'bg_color': '#E6CCFF'  # Purple for child cells 
    })
    
    transport_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter',
        'bg_color': '#FFD0A4'  # Orange for transport
    })
    
    # Create cell formats for each booking group color
    color_formats = []
    for color in color_palette:
        color_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'bg_color': color
        })
        color_formats.append(color_format)
    
    default_cell_format = workbook.add_format({
        'border': 1,
        'valign': 'vcenter'
    })
    
    # Write headers - use trip_type for the header
    headers = [
        'A/A', 'SOURCE', f"{tour_date} {trip_type} Tour", 'NAME', 'GEN', 'CONTACT', 'NAT', 
        'DATE of BIRTH', 'PASSPORT', 'TRSF', 'Tr.COST', 'BOOKING TOTAL', 
        'PRICE', 'PAYABLE', 'COMMENTS', 'Online Payed Total', '', '', ''
    ]
    
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)
    
    # Set column widths
    column_widths = [5, 12, 25, 25, 5, 15, 20, 15, 15, 10, 10, 15, 10, 10, 20, 20, 5, 5, 30]
    for col, width in enumerate(column_widths):
        worksheet.set_column(col, col, width)
    
    row = 1
    passenger_count = 1
    
    # Process each booking group
    for booking_idx, booking in enumerate(bookings):
        # Determine source code for output
        source_code = booking.get('source_code', "")
        source = booking.get('source', "")
        
        # Use source_code if available, otherwise use source
        if source_code:
            output_source = source_code
        elif source:
            output_source = source
        else:
            output_source = "Unknown"
        
        # Clean up source if it has "znb" prefix
        if "znb" in output_source.lower():
            output_source = output_source.lower().replace("znb", "").strip().upper()
        
        # Get color for this booking group
        color_idx = booking_idx % len(color_palette)
        booking_color_format = color_formats[color_idx]
        
        # Extract transport details from passengers or comments
        pickup_location = ""
        pickup_time = ""
        
        # First, check if any passenger has pickup location info
        for passenger in booking['passengers']:
            if passenger.get('pickup_location'):
                pickup_location = passenger['pickup_location']
                pickup_time = passenger.get('pickup_time', '')
                break
        
        # If not found in passengers, check comments
        if not pickup_location and booking.get('comments'):
            comments = booking['comments']
            # Look for pickup location in comments
            pickup_match = re.search(r'Pickup location:\s*([^\n]+?)\s*(?:Pickup time:|$)', comments)
            if pickup_match:
                pickup_location = pickup_match.group(1).strip()
            
            # Look for pickup time in comments
            time_match = re.search(r'Pickup time:\s*([^\n]+?)\s*(?:Edit|$)', comments)
            if time_match:
                pickup_time = time_match.group(1).strip()
        
        # Process each passenger in this booking
        for i, passenger in enumerate(booking['passengers']):
            # Determine price based on passenger type and source
            price = 40  # Default adult price
            if passenger['type'] == "Child":
                price = 20
            elif passenger['type'] == "Infant":
                price = 0
            elif output_source == "Letsbook":
                price = 0
            
            # Determine transport cost based on pickup location
            transport_location = ""
            transport_cost = ""
            if booking['transport'] == "Yes":
                transport_location = pickup_location
                if "Adamas" in transport_location:
                    transport_cost = "20"
                elif "Pollonia" in transport_location:
                    transport_cost = "30"
            
            # Clean total and due amounts
            total_value = booking['total'].replace('€', '').replace(',', '').strip()
            due_value = booking['due'].replace('€', '').replace(',', '').strip()
            
            # Write data with booking color for A/A column
            worksheet.write(row, 0, passenger_count, booking_color_format)  # A/A with booking color
            
            # Now write each cell with appropriate formatting
            for col in range(1, 19):  # Columns B to S (1-18)
                cell_value = ""
                cell_format = default_cell_format
                
                # Determine cell value based on column
                if col == 1:  # SOURCE
                    cell_value = output_source
                elif col == 2:  # Tour date and type
                    cell_value = f"{tour_date} {trip_type} Tour"
                elif col == 3:  # NAME
                    cell_value = passenger.get('full_name', '')
                elif col == 4:  # GEN
                    gender = passenger.get('gender', '')
                    if gender.upper() == 'MALE':
                        gender = 'M'
                    elif gender.upper() == 'FEMALE':
                        gender = 'F'
                    elif len(gender) > 1:
                        gender = gender[0].upper()
                    cell_value = gender
                elif col == 5:  # CONTACT
                    cell_value = passenger.get('phone', '')
                elif col == 6:  # NAT
                    cell_value = passenger.get('nationality', '')
                elif col == 7:  # DATE of BIRTH
                    dob = passenger.get('dob', '')
                    if dob:
                        dob = dob.replace('.', '/').replace('-', '/')
                    cell_value = dob
                elif col == 8:  # PASSPORT
                    cell_value = passenger.get('passport', '')
                elif col == 9:  # TRSF
                    cell_value = transport_location
                elif col == 10:  # Tr.COST
                    cell_value = transport_cost
                elif col == 11:  # BOOKING TOTAL
                    cell_value = total_value if i == 0 else ""
                elif col == 12:  # PRICE
                    cell_value = price
                elif col == 13:  # PAYABLE
                    cell_value = due_value if i == 0 else ""
                elif col == 14:  # COMMENTS
                    cell_value = booking.get('comments', '')
                elif col == 15:  # Online Payed Total
                    cell_value = ""
                elif col == 18:  # Email (last column)
                    cell_value = booking.get('email', '') if i == 0 else ""
                
                # IMPORTANT: Apply formats in the correct order
                # 1. First apply transport format if needed
                if booking['transport'] == "Yes" and 1 <= col <= 14:
                    cell_format = transport_format
                
                # 2. Then apply child format if passenger is a child (overrides transport for D-G)
                if passenger['type'] == "Child" and 3 <= col <= 6:
                    cell_format = child_cells_format
                
                # 3. Then apply infant format if passenger is an infant (overrides transport for D-G)
                if passenger['type'] == "Infant" and 3 <= col <= 6:
                    cell_format = infant_cells_format
                
                # Write the cell with the determined format
                worksheet.write(row, col, cell_value, cell_format)
            
            row += 1
            passenger_count += 1
    
    # Add a summary
    worksheet.write(row + 2, 0, f"Total Passengers: {passenger_count - 1}", workbook.add_format({'bold': True}))
    worksheet.write(row + 3, 0, f"Total Bookings: {len(bookings)}", workbook.add_format({'bold': True}))
    
    # Add legend for colors
    worksheet.write(row + 5, 0, "Color Legend:", workbook.add_format({'bold': True}))
    
    # Infant legend
    infant_legend_format = workbook.add_format({'bg_color': '#CCE5FF', 'border': 1})
    worksheet.write(row + 6, 0, "Infant", infant_legend_format)
    worksheet.write(row + 6, 1, "Cells D-G (NAME to NAT) colored light blue")
    
    # Child legend
    child_legend_format = workbook.add_format({'bg_color': '#E6CCFF', 'border': 1})
    worksheet.write(row + 7, 0, "Child", child_legend_format)
    worksheet.write(row + 7, 1, "Cells D-G (NAME to NAT) colored orange")
    
    # Transport legend
    transport_legend_format = workbook.add_format({'bg_color': '#FFD0A4', 'border': 1})
    worksheet.write(row + 8, 0, "Transport", transport_legend_format)
    worksheet.write(row + 8, 1, "Cells B-O (SOURCE to COMMENTS) colored purple")
    
    workbook.close()
    output.seek(0)
    return output

def extract_date_from_filename(filename):
    """Extract date from filename in month-day-year format and convert to day-month-year"""
    # Remove file extension
    filename_without_ext = os.path.splitext(filename)[0]
    
    # Try to find date pattern in filename (month-day-year)
    date_pattern = r'(\d{1,2})[-_](\d{1,2})[-_](\d{4})'
    match = re.search(date_pattern, filename_without_ext)
    
    if match:
        month = match.group(1).zfill(2)
        day = match.group(2).zfill(2)
        year = match.group(3)
        
        # Return in day-month-year format
        return f"{day}-{month}-{year}"
    else:
        # If no date found in filename, return today's date
        today = datetime.now()
        return today.strftime("%d-%m-%Y")

def analyze_transfers(bookings):
    """Analyze transfer pickup locations and count passengers per location"""
    transfer_counts = {}
    
    for booking in bookings:
        if booking.get('transport') == "Yes":
            # Extract pickup location from passengers or comments
            pickup_location = ""
            
            # First, check if any passenger has pickup location info
            for passenger in booking['passengers']:
                if passenger.get('pickup_location'):
                    pickup_location = passenger['pickup_location']
                    break
            
            # If not found in passengers, check comments
            if not pickup_location and booking.get('comments'):
                comments = booking['comments']
                pickup_match = re.search(r'Pickup location:\s*([^\n]+?)\s*(?:Pickup time:|$)', comments)
                if pickup_match:
                    pickup_location = pickup_match.group(1).strip()
            
            # Clean up the pickup location
            if pickup_location:
                # Standardize common pickup locations
                if 'adamas' in pickup_location.lower() or 'adamantas' in pickup_location.lower():
                    pickup_location = "Adamas Port"
                elif 'pollonia' in pickup_location.lower():
                    pickup_location = "Pollonia"
                elif 'airport' in pickup_location.lower():
                    pickup_location = "Airport"
                elif 'hotel' in pickup_location.lower():
                    # Try to extract hotel name
                    hotel_match = re.search(r'hotel\s+([^,\n]+)', pickup_location.lower())
                    if hotel_match:
                        hotel_name = hotel_match.group(1).strip().title()
                        pickup_location = f"Hotel {hotel_name}"
                    else:
                        pickup_location = "Hotel"
                
                # Count passengers for this location
                num_passengers = len(booking['passengers'])
                if pickup_location in transfer_counts:
                    transfer_counts[pickup_location] += num_passengers
                else:
                    transfer_counts[pickup_location] = num_passengers
    
    return transfer_counts

def get_top_5_sources(bookings):
    """Get top 5 sources by passenger count"""
    source_counts = {}
    
    for booking in bookings:
        # Determine source for this booking - USE THE SAME LOGIC AS IN create_output_file()
        source_code = booking.get('source_code', "")
        source = booking.get('source', "")
        
        # Use source_code if available, otherwise use source
        if source_code:
            output_source = source_code
        elif source:
            output_source = source
        else:
            output_source = "Unknown"
        
        # Clean up source if it has "znb" prefix
        if "znb" in output_source.lower():
            output_source = output_source.lower().replace("znb", "").strip().upper()
        
        # Count passengers for this source
        num_passengers = len(booking['passengers'])
        if output_source in source_counts:
            source_counts[output_source] += num_passengers
        else:
            source_counts[output_source] = num_passengers
    
    # Sort by count and get top 5
    sorted_sources = sorted(source_counts.items(), key=lambda x: x[1], reverse=True)
    
    # Return only top 5 (or less if fewer exist)
    top_5_dict = dict(sorted_sources[:5])
    
    return top_5_dict

def get_top_5_nationalities(bookings):
    """Get top 5 nationalities by passenger count"""
    nationality_counts = {}
    
    for booking in bookings:
        for passenger in booking['passengers']:
            # Get nationality exactly as it appears in passenger data
            nationality = passenger.get('nationality', '')
            
            # Clean it the same way as in the Excel output
            if nationality:
                nationality = nationality.strip()
            else:
                nationality = "Unknown"
            
            if nationality in nationality_counts:
                nationality_counts[nationality] += 1
            else:
                nationality_counts[nationality] = 1
    
    # Sort by count and get top 5
    sorted_nationalities = sorted(nationality_counts.items(), key=lambda x: x[1], reverse=True)
    
    # Return only top 5 (or less if fewer exist)
    top_5_dict = dict(sorted_nationalities[:5])
    
    return top_5_dict

def check_flagged_comments(bookings):
    """Check for flagged comments based on keywords"""
    # List of keywords to flag (can be expanded in the future)
    flagged_keywords = [
        'allergy', 'allergies', 'condition', 'birthday', 'pregnant',
        'medical', 'illness', 'sick', 'disability', 'wheelchair',
        'dietary', 'vegetarian', 'vegan', 'gluten', 'diabetic',
        'prescription', 'medicine', 'pregnancy', 'celebrate',
        'anniversary', 'special needs', 'health', 'injury'
    ]
    
    flagged_items = []
    passenger_count = 1  # To match the A/A numbering in Excel
    
    for booking in bookings:
        booking_comment = booking.get('comments', '').lower()
        
        # Check if booking comment contains any flagged keywords
        if booking_comment:
            found_keywords = []
            for keyword in flagged_keywords:
                if keyword in booking_comment:
                    found_keywords.append(keyword)
            
            if found_keywords:
                # Get the first passenger's name for reference
                first_passenger_name = ""
                if booking['passengers'] and booking['passengers'][0].get('full_name'):
                    first_passenger_name = booking['passengers'][0]['full_name']
                
                # Add each passenger in the booking (since the comment applies to all)
                for i, passenger in enumerate(booking['passengers']):
                    passenger_name = passenger.get('full_name', f'Passenger {passenger_count}')
                    flagged_items.append({
                        'passenger_number': passenger_count,
                        'passenger_name': passenger_name,
                        'booking_id': booking.get('booking_id', ''),
                        'comment': booking.get('comments', ''),
                        'keywords': found_keywords,
                        'first_in_booking': (i == 0)  # Mark if first passenger in booking
                    })
                    passenger_count += 1
        else:
            # No comment, just increment passenger count
            passenger_count += len(booking['passengers'])
    
    return flagged_items

def main():
    # Add logo at the top
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Try to load logo from various locations
        logo_paths = [
            "logo.png",
            "logo.jpg",
            "logo.jpeg",
            "images/logo.png",
            "static/logo.png",
            "assets/logo.png"
        ]
        
        logo_found = False
        for logo_path in logo_paths:
            try:
                if os.path.exists(logo_path):
                    logo = Image.open(logo_path)
                    st.image(logo, use_column_width=True)
                    logo_found = True
                    break
            except:
                continue
        
        if not logo_found:
            st.markdown("### ⚓ .csv list parser ⚓")
    
    st.markdown("Upload a manifest .CSV file, to extract the formatted passenger list")
    st.markdown("*don't forget to ***expand*** the booking info or else all will fail* 💀")
    
    # No need for sidebar input anymore since we extract from filename and CSV
    
    # File upload - now accepting CSV
    uploaded_file = st.file_uploader("Choose a manifest CSV file", type=['csv'])
    
    if uploaded_file is not None:
        # Extract date from filename
        extracted_date = extract_date_from_filename(uploaded_file.name)
        
        # Show file info
        st.info(f"Uploaded file: {uploaded_file.name}")
        st.info(f"Extracted date: {extracted_date}")
        
        # Parse the file
        with st.spinner("Parsing manifest CSV file..."):
            bookings, trip_type = parse_manifest_csv(uploaded_file, uploaded_file.name)
        
        # Show trip type
        st.info(f"Trip type: {trip_type} Tour")
        
        if bookings:
            # Calculate total passengers
            total_passengers = sum(len(b['passengers']) for b in bookings)
            
            st.success(f"✅ Successfully parsed {len(bookings)} bookings with {total_passengers} passengers")
            
            # Show statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Bookings", len(bookings))
            with col2:
                st.metric("Total Passengers", total_passengers)
            with col3:
                transport_count = sum(1 for b in bookings if b.get('transport') == "Yes")
                st.metric("Transport Requests", transport_count)
            with col4:
                # Count infants and children
                infant_count = 0
                child_count = 0
                for b in bookings:
                    for p in b['passengers']:
                        if p.get('type') == "Infant":
                            infant_count += 1
                        elif p.get('type') == "Child":
                            child_count += 1
                st.metric("Infants/Children", f"{infant_count}/{child_count}")
            
            # Create output file
            with st.spinner("Creating output file with enhanced formatting..."):
                output_file = create_output_file(bookings, extracted_date, trip_type)
            
            # Create output filename in format: Date (as day-month-year) + trip type
            output_filename = f"{extracted_date} {trip_type} Tour.xlsx"
            
            # Download button
            st.download_button(
                label=f"📥 Download {output_filename}",
                data=output_file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help=f"Download the Excel file for {extracted_date} {trip_type} Tour"
            )
            
            # Flagged Comments section
            flagged_comments = check_flagged_comments(bookings)
            
            if flagged_comments:
                st.subheader("🚩 Flagged Comments")
                st.info(f"Found {len(flagged_comments)} passenger(s) with flagged comments")
                
                # Group by booking to avoid duplicates
                seen_bookings = set()
                for item in flagged_comments:
                    if item['booking_id'] not in seen_bookings:
                        seen_bookings.add(item['booking_id'])
                        
                        # Get all passengers in this booking with flagged comment
                        booking_passengers = [i for i in flagged_comments if i['booking_id'] == item['booking_id']]
                        
                        if booking_passengers:
                            # Show booking once with all passengers
                            passenger_nums = sorted([p['passenger_number'] for p in booking_passengers])
                            passenger_names = [p['passenger_name'] for p in booking_passengers[:3]]  # Show first 3 names
                            
                            # Format passenger numbers
                            if len(passenger_nums) == 1:
                                passenger_ref = f"Person {passenger_nums[0]}"
                            elif len(passenger_nums) <= 3:
                                passenger_ref = f"Persons {', '.join(map(str, passenger_nums))}"
                            else:
                                passenger_ref = f"Persons {passenger_nums[0]}-{passenger_nums[-1]}"
                            
                            # Show names if available
                            name_info = ""
                            if any(p['passenger_name'] for p in booking_passengers[:3]):
                                name_info = f" ({', '.join([n for n in passenger_names if n])})"
                            
                            # Show the comment
                            st.warning(f"**{passenger_ref}{name_info}** informs about: {item['comment']}")
                            st.caption(f"*Keywords detected: {', '.join(item['keywords'])}*")
                            st.write("---")
            else:
                st.subheader("🚩 Flagged Comments")
                st.success("No flagged comments found.")
            
            # Show data insights
            st.subheader("📈 Data Insights")
            
            # First row: Transfers and Revenue
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Transfers:**")
                transfer_counts = analyze_transfers(bookings)
                
                if transfer_counts:
                    # Sort by number of passengers (descending)
                    sorted_transfers = sorted(transfer_counts.items(), key=lambda x: x[1], reverse=True)
                    
                    total_transfer_passengers = sum(transfer_counts.values())
                    st.write(f"Total passengers needing transfer: **{total_transfer_passengers}**")
                    st.write("")
                    
                    for location, count in sorted_transfers:
                        st.write(f"• **{location}**: {count} passenger(s)")
                else:
                    st.write("No transfers requested for this tour.")
            
            with col2:
                st.write("**Revenue Summary:**")
                total_revenue = 0
                total_due = 0
                for b in bookings:
                    try:
                        total_str = b.get('total', '').replace('€', '').replace(',', '').strip()
                        due_str = b.get('due', '').replace('€', '').replace(',', '').strip()
                        if total_str:
                            total_revenue += float(total_str)
                        if due_str:
                            total_due += float(due_str)
                    except ValueError:
                        pass
                st.write(f"• Total Revenue: €{total_revenue:,.2f}")
                st.write(f"• Amount Due: €{total_due:,.2f}")
                st.write(f"• Collected: €{total_revenue - total_due:,.2f}")
            
            # Second row: Pie Charts
            st.subheader("Passenger Demographics")
            
            col3, col4 = st.columns(2)
            
            with col3:
                st.write("**Top 5 Sources (by Passengers)**")
                source_data = get_top_5_sources(bookings)
                
                if source_data:
                    # Create pie chart
                    labels = list(source_data.keys())
                    values = list(source_data.values())
                    
                    fig1 = go.Figure(data=[go.Pie(
                        labels=labels,
                        values=values,
                        hole=0.3,
                        textinfo='label+percent',
                        textposition='inside',
                        insidetextorientation='radial'
                    )])
                    
                    fig1.update_layout(
                        showlegend=False,
                        margin=dict(t=0, b=0, l=0, r=0),
                        height=300
                    )
                    
                    st.plotly_chart(fig1, width='stretch')
                    
                    # Show exact numbers
                    st.write("**Exact Counts:**")
                    for source, count in source_data.items():
                        st.write(f"• {source}: {count} passenger(s)")
                else:
                    st.write("No source data available.")
            
            with col4:
                st.write("**Top 5 Nationalities (by Passengers)**")
                nationality_data = get_top_5_nationalities(bookings)
                
                if nationality_data:
                    # Create pie chart
                    labels = list(nationality_data.keys())
                    values = list(nationality_data.values())
                    
                    fig2 = go.Figure(data=[go.Pie(
                        labels=labels,
                        values=values,
                        hole=0.3,
                        textinfo='label+percent',
                        textposition='inside',
                        insidetextorientation='radial'
                    )])
                    
                    fig2.update_layout(
                        showlegend=False,
                        margin=dict(t=0, b=0, l=0, r=0),
                        height=300
                    )
                    
                    st.plotly_chart(fig2, width='stretch')
                    
                    # Show exact numbers
                    st.write("**Exact Counts:**")
                    for nationality, count in nationality_data.items():
                        st.write(f"• {nationality}: {count} passenger(s)")
                else:
                    st.write("No nationality data available.")

if __name__ == "__main__":
    main()