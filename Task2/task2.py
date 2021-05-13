import ClointFusion as cf
cf.OFF_semi_automatic_mode()

def read_coordinates(img):
    coordinates = cf.mouse_search_snip_return_coordinates_x_y(img,conf=0.8)
    return coordinates

def click_mouse(coordinates,single_double_triple):
    cf.mouse_click(x=coordinates[0],y=coordinates[1],single_double_triple=single_double_triple)
    print(f"Clicked on {coordinates} {single_double_triple} times")

def searchProduct_and_GoToCart(productName,img):
    cf.mouse_find_highlight_click(searchText='Search for products, brands and more')

    cf.message_counter_down_timer(strMsg=f'Searching for {productName}',start_value=3)

    cf.key_write_enter(strMsg=productName,delay=1,key='e')

    cf.message_counter_down_timer(strMsg=f'find {productName}',start_value=3)

    coordinates = read_coordinates(img)
    click_mouse(coordinates,'single')

    # Adding to Cart
#     cf.message_counter_down_timer(strMsg='add to Cart',start_value=3)
#     coordinates = read_coordinates(img='img/add_to_cart.png')
#     click_mouse(coordinates,'single')
    
    # go to cart
    cf.message_counter_down_timer(strMsg='Go to Cart',start_value=3)
    coordinates = read_coordinates(img='img/go_to_cart.png')
    click_mouse(coordinates,'single')

# to open browser
cf.window_show_desktop()

coordinates = read_coordinates(img='img/chrome.png')

click_mouse(coordinates,'double')

# now select new tab than it will open flipkart
cf.window_activate_and_maximize_windows()

cf.key_write_enter(strMsg='https://flipkart.com',key='e')

# Product-1
productName = 'APPLE Watch Series 3 GPS - 42 mm Space Grey Aluminium Case with Black Sport Band  (Black Strap, Regular)'
img = 'img/APPLE Watch Series 3 GPS - 42 mm Space Grey Aluminium Case with Black Sport Band  (Black Strap, Regular).png'
searchProduct_and_GoToCart(productName,img)

# Product-2
productName = 'APPLE Watch Series 3 42 mm Silver Aluminum White Sport Band'
img = 'img/APPLE Watch Series 3 42 mm Silver Aluminum White Sport Band.png'
searchProduct_and_GoToCart(productName,img)

# Place Order
cf.message_counter_down_timer(strMsg='Placing_Order',start_value=3)

coordinates = read_coordinates(img='img/place_order.png')
click_mouse(coordinates,'single')
coordinates = read_coordinates(img='img/continue.png')
click_mouse(coordinates,'single')

# card detail
cvv = cf.gui_get_any_input_from_user(msgForUser='Enter cvv',password=True)

cf.key_write_enter(cvv)

# closing
coordinates = read_coordinates(img='img/close.png')
click_mouse(coordinates,'single')

cf.message_counter_down_timer(strMsg='Exit Browser')