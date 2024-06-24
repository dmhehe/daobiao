import json

import weakref

import simui.simtimer

class SimList(list):
    def _check_save_json(self, attr_name=None):
        pass
    
    def parent_check_save_json(self):
        retrieved_obj = self.ref()
        if retrieved_obj:
            retrieved_obj._check_save_json()
    
    def set_parent(self, parent):
        self.ref = weakref.ref(parent)
        
    def set_ui_obj(self, ui_obj):
        self.ui_obj_ref = weakref.ref(ui_obj)
        
    def set_callback(self, callback):
        self.callback = callback
        
    def trigger_callback(self):
        if hasattr(self, "callback") and self.callback:
            self.callback()
        
        if self.__dict__.get("__showing", False):
            simui.simtimer.add_timer("SimList_refesh_ui", 10, self.ui_obj_ref().refesh_ui)
        
        self.parent_check_save_json()
        

    def __setitem__(self, index, value):
        super(SimList, self).__setitem__(index, value)
        self.trigger_callback()

    def __delitem__(self, index):
        super(SimList, self).__delitem__(index)
        self.trigger_callback()

    def append(self, value):
        super(SimList, self).append(value)
        self.trigger_callback()

    def extend(self, iterable):
        super(SimList, self).extend(iterable)
        self.trigger_callback()

    def remove(self, value):
        super(SimList, self).remove(value)
        self.trigger_callback()

    def pop(self, index=-1):
        value = super(SimList, self).pop(index)
        self.trigger_callback()
        return value
    

    
    