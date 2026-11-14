import sys
if __name__ == '__main__':
    if '--self-test' in sys.argv:
        from pathlib import Path
        from datetime import timedelta
        import tempfile,uuid
        from returns_app.server import create_app
        from returns_app.domain import ReceiptInput,now
        with tempfile.TemporaryDirectory() as folder:
            app=create_app(Path(folder),'s'*32,'a'*32,'SelfTest123')
            store=app.state.store
            receipt=store.create_receipt(ReceiptInput(request_id=str(uuid.uuid4()),operator='مطيع',customer_code=store.search('customers')['items'][0]['code'],items=[{'product_code':store.search('products')['items'][0]['code'],'cartons':1,'expiry':(now().date()+timedelta(days=60)).isoformat()}]))
            assert store.receipt(receipt['serial'])[1].startswith(b'%PDF')
        print('Server self-test passed.')
        raise SystemExit(0)
    from returns_app.server import main
    main()
