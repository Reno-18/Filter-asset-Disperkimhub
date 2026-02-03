"""
AsetFilter - Flask Forms
"""
from flask_wtf import FlaskForm
from flask_wtf.file import FileField, FileAllowed, FileRequired
from wtforms import StringField, FloatField, SelectField, SelectMultipleField, SubmitField
from wtforms.validators import Optional, NumberRange


class UploadForm(FlaskForm):
    """Form for uploading Excel files"""
    file = FileField('Excel File', validators=[
        FileRequired(message='Please select a file'),
        FileAllowed(['xls', 'xlsx'], message='Only Excel files (.xls, .xlsx) are allowed')
    ])
    submit = SubmitField('Upload & Process')


class FilterForm(FlaskForm):
    """Form for filtering assets"""
    nama_asset = StringField('Nama Asset', validators=[Optional()])
    kecamatan = SelectField('Kecamatan', choices=[], validators=[Optional()])
    min_luas = FloatField('Min Luas (m²)', validators=[Optional(), NumberRange(min=0)])
    max_luas = FloatField('Max Luas (m²)', validators=[Optional(), NumberRange(min=0)])
    status = SelectMultipleField('Status', choices=[], validators=[Optional()])
    submit = SubmitField('Apply Filters')
    reset = SubmitField('Reset Filters')
