import weakref


from tableaudocumentapi import Datasource, xfile
from tableaudocumentapi.xfile import xml_open
class Preferences(object):

    """
    A class for writing Tableau preference files.

    """

    def __init__(self, filename):
        """
        Constructor.

        """

        self._filename = filename

        self._preferenceTree = xml_open(self._filename, self.__class__.__name__.lower())

        self._preferenceRoot = self._preferenceTree.getroot()

        self._palettes = self._get_palettes(self._preferenceRoot)


    @property
    def palettes(self):
        return self._palettes

    @staticmethod
    def _get_palettes(xml_root):
        colorPalette = []
        for child in xml_root.iter('color-palette'):
            # print(child.attrib)
            colorPalette.append(child)
        return colorPalette




class Workbook(object):
    """A class for writing Tableau workbook files."""

    def __init__(self, filename):
        """Open the workbook at `filename`. This will handle packaged and unpacked
        workbook files automatically. This will also parse Data Sources and Worksheets
        for access.
        """

        self._filename = filename

        self._workbookTree = xml_open(self._filename, 'workbook')

        self._workbookRoot = self._workbookTree.getroot()
        # prepare our datasource objects
        self._datasources = self._prepare_datasources(
            self._workbookRoot)  # self.workbookRoot.find('datasources')

        self._datasource_index = self._prepare_datasource_index(self._datasources)

        self._worksheets = self._prepare_worksheets(
            self._workbookRoot, self._datasource_index
        )
        #    
        self._dashboards = self._prepare_dashboards(
            self._workbookRoot
        )
        
        #
        self._palette = self._list_palette(
            self._workbookRoot
        )

        #
        self._colorColumns = self._list_color_columns(
            self._workbookRoot
        )
        
        #
        self._encodedcolorColumns = self._list_encoded_color_columns(
            self._workbookRoot
        )
        
        #
        self._mapColorCols = self._list_map_color_columns(
            self._workbookRoot
        )

    @property
    def datasources(self):
        return self._datasources

    @property
    def worksheets(self):
        return self._worksheets
    
    
    @property
    def dashboards(self):
        return self._dashboards

    @property
    def palette(self):
        return self._palette

    @property
    def color_columns(self):
        return self._colorColumns

    @property
    def encoded_color_columns(self):
        return self._encodedcolorColumns

    @property
    def map_color_columns(self):
        return self._mapColorCols



    @property
    def filename(self):
        return self._filename

    def save(self):
        """
        Call finalization code and save file.
        Args:
            None.
        Returns:
            Nothing.
        """

        # save the file
        xfile._save_file(self._filename, self._workbookTree)

    def save_as(self, new_filename):
        """
        Save our file with the name provided.
        Args:
            new_filename:  New name for the workbook file. String.
        Returns:
            Nothing.
        """
        xfile._save_file(
            self._filename, self._workbookTree, new_filename)

    @staticmethod
    def _prepare_datasource_index(datasources):
        retval = weakref.WeakValueDictionary()
        for datasource in datasources:
            retval[datasource.name] = datasource

        return retval

    @staticmethod
    def _prepare_datasources(xml_root):
        datasources = []

        # loop through our datasources and append
        datasource_elements = xml_root.find('datasources')
        if datasource_elements is None:
            return []

        for datasource in datasource_elements:
            ds = Datasource(datasource)
            datasources.append(ds)

        return datasources

    @staticmethod
    def _prepare_worksheets(xml_root, ds_index):
        worksheets = []
        worksheets_element = xml_root.find('.//worksheets')
        if worksheets_element is None:
            return worksheets

        for worksheet_element in worksheets_element:
            worksheet_name = worksheet_element.attrib['name']
            worksheets.append(worksheet_name)  # TODO: A real worksheet object, for now, only name

            dependencies = worksheet_element.findall('.//datasource-dependencies')

            for dependency in dependencies:
                datasource_name = dependency.attrib['datasource']
                datasource = ds_index[datasource_name]
                for column in dependency.findall('.//column'):
                    column_name = column.attrib['name']
                    if column_name in datasource.fields:
                        datasource.fields[column_name].add_used_in(worksheet_name)

        return worksheets
    
    
    
        ###############################################################################
    #
    # New Changes - A method to get all Dashboard and the corresponding sheet details
    #
    ###############################################################################

    @staticmethod
    def _prepare_dashboards(xml_root):
        # tree = ET.parse(xml_root)
        # root = tree.getroot()
        dashboards = {}
        for child in xml_root.iter('dashboard'):
            worksheets = []
            try:
                if bool(child.attrib) and child.attrib['name'] and not child.attrib['type']:
                    i = 0
            # print(child.attrib)
            except KeyError:
                if bool(child.attrib) and child.attrib['name']:
                    i = 0
                        # print(child.attrib['name'])
                        # print(ET.dump(child))
                    for gc in child.iter('zone'):
                        # print(gc.attrib)
                        if bool('name' in gc.attrib) and gc.attrib['name'] not in worksheets:
                            # DWorksheet[child.attrib['name']].append(worksheets)
                            worksheets.append(gc.attrib['name'])
                            i = i+1
                    if i == 0:
                        # NullDashboard.append(child.attrib['name'])
                        pass
                    dashboards.update({child.attrib['name']: worksheets})
                    # Dashboards.update({child:worksheets})
                    # TODO Can we share the Dashboard Element in a Dict ..comment above line

        return dashboards

    def _remove_dashboards(self, name):
        xml_root = self._workbookRoot
        delete = []
        for child in xml_root.iter('dashboard'):
            # print(child.attrib)
            try:
                if child.attrib['name'] in name:
                    delete.append(child)
            except KeyError:
                pass
        for child in xml_root.iter('dashboards'):
            for dashboard in delete:
                child.remove(dashboard)

        story_point_elements = []

        for child in xml_root.iter('story-point'):
            try:
                if child.attrib['captured-sheet'] == name:
                    story_point_elements.append(child)
            except KeyError:
                pass
        # print(story_point_elements)

        for child in xml_root.iter('story-points'):
            for element in story_point_elements:
                child.remove(element)

        # Remove blank dashboard tag

        for child in xml_root.iter('dashboards'):
            if len(list(child)) == 0:
                # print(child)
                xml_root.remove(child)
        xfile._save_file(self._filename, self._workbookTree)

    
    @staticmethod
    def _list_palette(xml_root):
        color_palette = []
        for child in xml_root.iter('color-palette'):
            try:
                color_palette.append(child.attrib['name'])
            except KeyError:
                pass
        return color_palette

    @staticmethod
    def _list_color_columns(xml_root):
        color_columns = []
        for child in xml_root.findall('.//encodings/'):
            if child.tag == 'color':
                column = child.attrib['column'].split('].', 1)[1]
                color_columns.append(column)
        return color_columns

    @staticmethod
    def _list_encoded_color_columns(xml_root):
        actual_color_columns = []
        for child in xml_root.iter('style-rule'):
            if child.attrib['element'] == 'mark':
                for c in child.iter('encoding'):
                    actual_color_columns.append(c.attrib['field'])
        return actual_color_columns

    @staticmethod
    def _list_map_color_columns(xml_root):
        column_color_mapping = {}
        for parent in xml_root.iter('style-rule'):

            for child in parent.iter('encoding'):
                # print(child.attrib)
                try:
                    palette = child.attrib['palette']
                except KeyError:
                    palette = ""
                column_color_mapping.update({str(child.attrib['field'])[6:-4]: palette})
        return column_color_mapping

    # @palette.setter

    def _set_palettes(self, preference_path, palette):
        preference = Preferences(preference_path)

        for i in preference.palettes:

            if i.attrib['name'] == palette:
                elem = i
        # print(sourcePF.palettes[0].attrib)
        xml_root = self._workbookRoot
        # print(xml_root)
        for parent in xml_root.findall('.//preferences/..'):
            # Find each color element
            for element in parent.findall('preferences'):
                # Define what the custom BofA color palette is
                # print(element.attrib)
                element.append(elem)
        # print(xml_root)
        xfile._save_file(self._filename, self._workbookTree)

    def _set_map_color_column_instances(self, columns, palette):
        # Assign the color column the palette within encodings tag
        xml_root = self._workbookRoot

        for parent in xml_root.findall('.datasources/'):
            if parent.attrib['name'] != 'Parameters':

                style_rule = False
                for child in parent.iter('style-rule'):
                    style_rule = True

                if style_rule is False:
                    list_parent = []
                    for child in parent:
                        list_parent.append(child)
                    # print(list_parent)

                    for child in parent.iter('layout'):
                        for i in range(len(list_parent)):
                            if list_parent[i] == child:
                                # print(child.attrib)
                                break
                            i += 1
                    append_style = ET.fromstring("""<style>
                                <style-rule element='mark'>
                                 
                                 </style-rule>
                                 </style>	""")
                    if i > 0:
                        parent.insert(i+1, append_style)

                for child in parent.iter('style-rule'):
                    # style_rule = True
                    # 3rd Change as per mail
                    # Finds the parent of each of the 'encodings' element
                    # print(child.attrib)
                    if child.attrib['element'] == 'mark':
                        # print(child)
                        for column in columns:
                            string_value = """<encoding attr='color' field='{}' palette='{}' type='palette'>
                                                                  </encoding>""".format(column, palette)
                            child.insert(0, ET.fromstring(string_value))

        xfile._save_file(self._filename, self._workbookTree)


    def _set_color_column_instances(self, columns):
        i=0
        xml_root = self._workbookRoot
        for parent in xml_root.findall('.datasources/'):
            if parent.attrib['name'] != 'Parameters':
                # Finds the parent of each of the 'encodings' element
                # 2nd change as per mail
                list_parent = []
                for child in parent:
                    list_parent.append(child)

                # print(ListParent)
                index = 0
                for child in parent.iter('_.fcp.ObjectModelTableType.true...column'):
                    for i in range(len(list_parent)):
                        if list_parent[i] == child:
                            index = i
                            # print(child.attrib)
                            break
                        i += 1
                if index == 0:
                    for child in parent.iter('column'):
                        for i in range(len(list_parent)):
                            if list_parent[i] == child:
                                index = i
                                # print(child.attrib)
                                # break
                            i += 1
                # print(i)
                key = {"sum": "Sum", "yr": "Year", "none": "None"}

                for column in columns:
                    try:
                        if column.split(":")[0][1:] == 'yr':
                            custom_add = ET.fromstring("""<column-instance column="[{}]" derivation="{}" name="{}" 
                            pivot="key" type="nominal" />""".format(column.split(":")[1], key[column.split(":")[0][1:]],
                                                                    column[:-3]+"ok]"))
                        else:
                            custom_add = ET.fromstring("""<column-instance column="[{}]" derivation="{}" name="{}" 
                            pivot="key" type="nominal" />""".format(column.split(":")[1], key[column.split(":")[0][1:]],
                                                                    column[:-3]+"nk]"))
                    except KeyError:
                        custom_add = ET.fromstring("""<column-instance column="[{}]" derivation="{}" name="{}" 
                        pivot="key" type="nominal" />""".format(column.split(":")[1], 'None',
                                                                column[:-3]+"nk]"))

                    if index > 0:
                        parent.insert(index+1, custom_add)

        xfile._save_file(self._filename, self._workbookTree)

    def _set_measure_map_color_column_instances(self, palette):
        # Assign the color column the palette within encodings tag
        xml_root = self._workbookRoot

        preference = Preferences('Preferences/Preferences.tps')
        for i in preference.palettes:
            if i.attrib['name'] == palette:
                elem = i
        # print(sourcePF.palettes[0].attrib)

        # print(xml_root)
        for parent in xml_root.findall('.//preferences/..'):
            # Find each color element
            for element in parent.findall('preferences'):
               element.append(elem)

        color_columns = []
        measure_color_columns = []
        for child in xml_root.findall('.//encodings/'):
            if child.tag == 'color':
                column = child.attrib['column'].split('.', 1)[1]
                if not "none" in child.attrib['column']:
                    measure_color_columns.append(column)
                else:
                    color_columns.append(column)
        # print(color_columns, measure_color_columns)

        worksheet_elems = []
        for parent in xml_root.iter('worksheet'):
            # print(parent.attrib)
            for child in parent.iter('color'):
                if "qk" in child.attrib['column']:
                    worksheet_elems.append(parent)

        # print(worksheet_elems)
        ds_name = []
        for elem in worksheet_elems:
            for parent in elem.iter('datasource'):
                ds_name.append(parent.attrib['name'])

        if len(measure_color_columns)>0 and len(ds_name)>0:
            string_val = """<style>
                      <style-rule element='mark'>
                        <encoding attr='color' field='[{}].{}' palette='{}' type='interpolated' />
                      </style-rule>
                    </style>""".format(ds_name[0], measure_color_columns[0], palette)
            append_style = ET.fromstring(string_val)

            for elem in worksheet_elems:
                # print(0, elem)
                for parent in elem.findall('table'):
                    # print(1, parent)
                    for child in parent:
                        if child.tag == 'style':
                            parent.remove(child)
                    # parent.remove(ET.fromstring("<style />"))
                    parent.insert(1, append_style)
        xfile._save_file(self._filename, self._workbookTree)