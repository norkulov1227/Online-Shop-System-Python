from shop_online_store import Shop


store = Shop('data.xlsx', 'status.xlsx')
store.read_xlsx()
store.read_status_xlsx()
store.main()


