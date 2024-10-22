# DTE
Traceback (most recent call last):
  File "C:\Users\siyiliu2\Desktop\cnn on defect\v1\label studio\U-2-Net\train_custom.py", line 74, in <module>
    main()
  File "C:\Users\siyiliu2\Desktop\cnn on defect\v1\label studio\U-2-Net\train_custom.py", line 30, in main
    dataloader = DataLoader(dataset, batch_size=8, shuffle=True, num_workers=4)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\site-packages\torch\utils\data\dataloader.py", line 351, in __init__
    sampler = RandomSampler(dataset, generator=generator)  # type: ignore[arg-type]
  File "C:\Users\siyiliu2\.conda\envs\image\lib\site-packages\torch\utils\data\sampler.py", line 144, in __init__
    raise ValueError(f"num_samples should be a positive integer value, but got num_samples={self.num_samples}")
ValueError: num_samples should be a positive integer value, but got num_samples=0
