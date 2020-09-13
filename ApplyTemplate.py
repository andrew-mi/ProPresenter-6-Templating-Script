import pandas #for opening .xlsx files; additionally requries xlrd installed
from lxml import etree #for opening .pro6 files (as xml)
from os import sys
import base64
import copy

TemplateXml = etree.parse('SampleTemplate.pro6') #Open pro6 document
for Slide in TemplateXml.findall('.//RVDisplaySlide'): #iterate to find template slide
    SlideLabel = Slide.attrib['label']
    if isinstance(SlideLabel, str) and SlideLabel.lower() == 'template':
        #Found first slide marked as template
        SlideArray = Slide.find('...') #Gets Parent by XPath
        SlideTemplateGroup = SlideArray.find('...') #ditto
        SlideTemplateGroup.attrib['name'] = 'Template'
        GroupsArray = SlideTemplateGroup.find('...')
        break
else:
    print('Cannot find template. Please label the slide as \'template\' in ProPresenter and try again.')
    sys.exit()

CurrentGroupLabel = ''

for row in pandas.read_excel('SampleData.xlsx').iterrows():
    Person = row[1]
    NewSlide = copy.deepcopy(Slide)

    #Make Slide
    for TextElementString in NewSlide.findall('.//NSString'):
        rvXMLIvarName = TextElementString.attrib['rvXMLIvarName']
        if isinstance(rvXMLIvarName, str) and rvXMLIvarName in ['PlainText','RTFData','WinFlowData']:
            DecodedString = base64.b64decode(TextElementString.text).decode('ascii')

            #Find and replace Name
            key = 'Name'
            value = Person['Name']
            DecodedString = DecodedString.replace(r'${' + key + r'$}', value).replace(r'$\\{' + key + r'$\\}', value).replace(r'{',r"\\{").replace(r'}',r"\\}")
            TextElementString.text = base64.b64encode(DecodedString.encode('ascii'))

            #Find and replace Description
            key = 'Description'
            value = Person['Description']
            if isinstance(value, str):
                DecodedString = DecodedString.replace(r'${' + key + r'$}', value).replace(r'$\\{' + key + r'$\\}', value).replace(r'{',r"\\{").replace(r'}',r"\\}")
                TextElementString.text = base64.b64encode(DecodedString.encode('ascii'))    
    #Replace Image
    if isinstance(Person['Image'], str):
        for ImageElement in NewSlide.findall('.//RVImageElement'):
            ImageElement.attrib['source'] = Person['Image']

    NewSlide.attrib['label'] = Person['Name']
    #NewSlide.attrib['notes'] = Person['Name']
    NewSlide.attrib['highlightColor'] = '0 0 0 1'

    #Make Group
    if not CurrentGroupLabel == Person['Label']:
        CurrentGroupLabel = Person['Label']
        NewGroup = etree.SubElement(GroupsArray, 'RVSlideGrouping')
        NewGroup.attrib['name'] = CurrentGroupLabel
        NewGroup.attrib['color'] = '%f %f %f 1' % {'C': (0.92549,0.03529,0.15294), 'F':(0.51373,0.18431,0.65491), 'L':(0.96471,0.65491,0.01569), 'T':(0.15294,0.33333,0.59216)}[CurrentGroupLabel[0]]
        NewGroup.attrib['uuid'] = 'CF56EA98-DC4C-4BFE-A336-AE4E5A893F8E'
        CurrentGroupSlideArray = etree.SubElement(NewGroup, 'array')
        CurrentGroupSlideArray.attrib['rvXMLIvarName'] = 'slides'

    CurrentGroupSlideArray.append(NewSlide) #Put newly constructed side in the appropriate group

etree.ElementTree(TemplateXml.getroot()).write('Output.pro6') #Save output