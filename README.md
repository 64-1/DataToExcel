# DataToExcel

Traceback (most recent call last):
  File "<string>", line 1, in <module>
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\spawn.py", line 116, in spawn_main
    exitcode = _main(fd, parent_sentinel)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\spawn.py", line 125, in _main
    prepare(preparation_data)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\spawn.py", line 236, in prepare
    _fixup_main_from_path(data['init_main_from_path'])
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\spawn.py", line 287, in _fixup_main_from_path
    main_content = runpy.run_path(main_path,
  File "C:\Users\siyiliu2\.conda\envs\image\lib\runpy.py", line 288, in run_path
    return _run_module_code(code, init_globals, run_name,
  File "C:\Users\siyiliu2\.conda\envs\image\lib\runpy.py", line 97, in _run_module_code
    _run_code(code, mod_globals, init_globals,
  File "C:\Users\siyiliu2\.conda\envs\image\lib\runpy.py", line 87, in _run_code
    exec(code, run_globals)
  File "C:\Users\siyiliu2\Desktop\cnn on defect\v1\label studio\U-2-Net\train_custom.py", line 44, in <module>
    for i, data in enumerate(dataloader):
  File "C:\Users\siyiliu2\.conda\envs\image\lib\site-packages\torch\utils\data\dataloader.py", line 440, in __iter__
    return self._get_iterator()
  File "C:\Users\siyiliu2\.conda\envs\image\lib\site-packages\torch\utils\data\dataloader.py", line 388, in _get_iterator
    return _MultiProcessingDataLoaderIter(self)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\site-packages\torch\utils\data\dataloader.py", line 1038, in __init__
    w.start()
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\process.py", line 121, in start
    self._popen = self._Popen(self)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\context.py", line 224, in _Popen
    return _default_context.get_context().Process._Popen(process_obj)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\context.py", line 327, in _Popen
    return Popen(process_obj)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\popen_spawn_win32.py", line 45, in __init__
    prep_data = spawn.get_preparation_data(process_obj._name)
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\spawn.py", line 154, in get_preparation_data
    _check_not_importing_main()
  File "C:\Users\siyiliu2\.conda\envs\image\lib\multiprocessing\spawn.py", line 134, in _check_not_importing_main
    raise RuntimeError('''
RuntimeError:
        An attempt has been made to start a new process before the
        current process has finished its bootstrapping phase.

        This probably means that you are not using fork to start your
        child processes and you have forgotten to use the proper idiom
        in the main module:

            if __name__ == '__main__':
                freeze_support()
                ...

        The "freeze_support()" line can be omitted if the program
        is not going to be frozen to produce an executable.
