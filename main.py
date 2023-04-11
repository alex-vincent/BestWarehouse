import pandas as pd
from haversine import haversine
from openpyxl import Workbook
from geopy.geocoders import Nominatim

WAREHOUSE_LOCATIONS = {
    "SDR Distribution": (43.720850, -79.661280),
    "SDR Distribution - Calgary": (51.158150, -114.000840),
    "Office": (47.603230, -122.330276),
    # Add more warehouses as needed
}


def read_csv_file(file_name):
    return pd.read_csv(file_name)


def haversine_distance(lat1, lon1, lat2, lon2):
    return haversine((lat1, lon1), (lat2, lon2), unit="km")


def get_coordinates_from_address(address):
    geolocator = Nominatim(user_agent="warehouse_allocator", timeout=5)
    location = geolocator.geocode(address)
    if location:
        print(f"Order with address: {address} has been geocoded sucessfully.")
        return (location.latitude, location.longitude)
    else:
        print(f"ERROR: Could not find coordinates for {address}")
        return None


def update_inventory_level(inventory_df, warehouse_id, sku, change):
    inventory_df.loc[
        (inventory_df["warehouse_id"] == warehouse_id) & (inventory_df["sku"] == sku),
        "inventory_level",
    ] += change


def allocate_orders(inventory_df, orders_df):
    allocations = []
    unallocated_orders = []

    for order_id, order_group in orders_df.groupby("Order ID"):
        sorted_warehouses = sorted(
            [
                (
                    warehouse_id,
                    haversine_distance(
                        order_group.iloc[0]["lat"],
                        order_group.iloc[0]["lon"],
                        coords[0],
                        coords[1],
                    ),
                )
                for warehouse_id, coords in WAREHOUSE_LOCATIONS.items()
            ],
            key=lambda x: x[1],
        )

        # Try to allocate the entire order to a single warehouse first
        allocated = False
        for warehouse_id, distance in sorted_warehouses:
            can_allocate_all = True
            for index, order in order_group.iterrows():
                required_sku = order["SKU"]
                required_quantity = order["Quantity"]

                try:
                    available_quantity = inventory_df.loc[
                        (inventory_df["warehouse_id"] == warehouse_id)
                        & (inventory_df["sku"] == required_sku),
                        "inventory_level",
                    ].values[0]
                except IndexError:
                    available_quantity = 0

                if available_quantity < 0:
                    available_quantity = 0

                if available_quantity < required_quantity:
                    can_allocate_all = False
                    break

            if can_allocate_all:
                for index, order in order_group.iterrows():
                    required_sku = order["SKU"]
                    required_quantity = order["Quantity"]
                    allocations.append(
                        (
                            order["Order ID"],
                            order["SKU"],
                            warehouse_id,
                            distance,
                            required_quantity,
                            False,
                        )
                    )
                    update_inventory_level(
                        inventory_df, warehouse_id, required_sku, -required_quantity
                    )
                allocated = True
                break

        # If the entire order can't be allocated to a single warehouse, try to split it
        if not allocated:
            for index, order in order_group.iterrows():
                required_sku = order["SKU"]
                required_quantity = order["Quantity"]

                allocated_quantity = 0

                for warehouse_id, distance in sorted_warehouses:
                    try:
                        available_quantity = inventory_df.loc[
                            (inventory_df["warehouse_id"] == warehouse_id)
                            & (inventory_df["sku"] == required_sku),
                            "inventory_level",
                        ].values[0]
                    except IndexError:
                        available_quantity = 0

                    if available_quantity < 0:
                        available_quantity = 0

                    if available_quantity >= required_quantity - allocated_quantity:
                        allocations.append(
                            (
                                order["Order ID"],
                                order["SKU"],
                                warehouse_id,
                                distance,
                                required_quantity - allocated_quantity,
                                True,
                            )
                        )
                        update_inventory_level(
                            inventory_df,
                            warehouse_id,
                            required_sku,
                            -(required_quantity - allocated_quantity),
                        )
                        allocated_quantity = required_quantity
                    else:
                        allocations.append(
                            (
                                order["Order ID"],
                                order["SKU"],
                                warehouse_id,
                                distance,
                                available_quantity,
                                True,
                            )
                        )
                        update_inventory_level(
                            inventory_df,
                            warehouse_id,
                            required_sku,
                            -available_quantity,
                        )
                        allocated_quantity += available_quantity

                    if allocated_quantity >= required_quantity:
                        break

                # Add unallocated orders to the list
                if allocated_quantity < required_quantity:
                    unallocated_orders.append(
                        (
                            order["Order ID"],
                            order["SKU"],
                            required_quantity - allocated_quantity,
                        )
                    )

    # Remove allocations with 0 quantity
    allocations = [allocation for allocation in allocations if allocation[4] > 0]

    return allocations, unallocated_orders


def create_excel_report(allocations):
    wb = Workbook()
    ws = wb.active
    ws.title = "Allocations"

    ws.cell(row=1, column=1, value="Order ID")
    ws.cell(row=1, column=2, value="SKU")
    ws.cell(row=1, column=3, value="Warehouse ID")
    ws.cell(row=1, column=4, value="Distance (km)")
    ws.cell(row=1, column=5, value="Allocated Quantity")
    ws.cell(row=1, column=6, value="Split Order")

    for idx, allocation in enumerate(allocations, 2):
        ws.cell(row=idx, column=1, value=allocation[0])
        ws.cell(row=idx, column=2, value=allocation[1])
        ws.cell(row=idx, column=3, value=allocation[2])
        ws.cell(row=idx, column=4, value=allocation[3])
        ws.cell(row=idx, column=5, value=allocation[4])
        ws.cell(row=idx, column=6, value=allocation[5])

    wb.save("allocation_report.xlsx")


def main():
    inventory_file = "inventory.csv"
    orders_file = "orders.csv"

    print("Reading inventory and orders data...")
    inventory_df = read_csv_file(inventory_file)
    orders_df = read_csv_file(orders_file)

    # Add latitude and longitude columns to the orders DataFrame
    orders_df["lat"] = 0.0
    orders_df["lon"] = 0.0
    for idx, order in orders_df.iterrows():
        address = f"{order['City']}, {order['Prov/ State']}, {order['Country Code']}"
        coordinates = get_coordinates_from_address(address)
        if coordinates:
            orders_df.at[idx, "lat"] = coordinates[0]
            orders_df.at[idx, "lon"] = coordinates[1]

    print("Allocating orders...")
    allocations, unallocated_orders = allocate_orders(inventory_df, orders_df)

    print("Creating Excel report...")
    create_excel_report(allocations)
    print("Allocation report generated: allocation_report.xlsx")

    if unallocated_orders:
        print("\nUnallocated Orders:")
        print("Order ID | SKU | Quantity")
        for unallocated_order in unallocated_orders:
            print(
                f"{unallocated_order[0]} | {unallocated_order[1]} | {unallocated_order[2]}"
            )
    else:
        print("\nAll orders have been allocated.")


if __name__ == "__main__":
    main()
