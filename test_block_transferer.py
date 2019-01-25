from dimensions_tools import copy_mdd_ddf_data, BlockTransferer

copy_mdd_ddf_data('test\\block_transferer\\clean', 'test\\block_transferer\\test')

bt = BlockTransferer('test\\block_transferer\\test.mdd', 'test\\block_transferer\\test.ddf', 'Test')
bt.update_mdd()
bt.update_ddf()