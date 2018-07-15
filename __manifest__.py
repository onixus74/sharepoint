# -*- coding: utf-8 -*-
{
    'name': "SharePoint",

   'summary': """SharePoint Odoo Integration""",

    'description': """
SharePoint-Odoo-Connector connects odoo with SharePoint.
It provides mutiple attachments in odoo.
When a file is uploaded in odoo it will also be uploaded at SharePoint.
In order to do so we just need Microsoft SharePoint URL,Username and Password.
Thers is an optional directory-name field which tells whether file should be placed in root directory or inside a directory.
	Overview
		Add SharePoint credential by using Setting under Configuration menu.
		URL,Username,password are mandatory fields and they required valid information to upload file at SharePoint.
		Add files using Upload File under Document menu.
		Directory name is kept optional if its value is empty then file will be uploaded to root path of alfresco repository otherwise it will create directory with <directory-name>
    """,

    'author': "Techscope",
    'website': "http://www.xyz.com",

    'category': 'Document Management',
    'version': '0.1',
    'depends': ['sale'],

    'data': [
        'views/templates.xml',
        'security/ir.model.access.csv',
    ],
    'demo': [],
}

