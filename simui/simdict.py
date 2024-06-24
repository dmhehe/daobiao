import json

import weakref

import simui.simtimer

class SimDict(dict):
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
            simui.simtimer.add_timer("SimDict_refesh_ui", 10, self.ui_obj_ref().refesh_ui)

        self.parent_check_save_json()

    def __setitem__(self, key, value):
        super(SimDict, self).__setitem__(key, value)
        self.trigger_callback()

    def __delitem__(self, key):
        super(SimDict, self).__delitem__(key)
        self.trigger_callback()

    def update(self, *args, **kwargs):
        super(SimDict, self).update(*args, **kwargs)
        self.trigger_callback()

    def pop(self, key, default=None):
        value = super(SimDict, self).pop(key, default)
        self.trigger_callback()
        return value