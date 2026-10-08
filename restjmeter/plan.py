"""Generate a standalone JMeter plan with CSV results and per-sample logging."""
import base64
from xml.etree import ElementTree as ET

from .workbook import xml_bytes


def literal(value):
    """Do not allow spreadsheet text to execute JMeter functions.

    JMeter does not unescape a string unless it contains an actual function.
    A fixed Groovy expression safely decodes data containing ${...}; base64
    cannot add Groovy code or function-argument delimiters.
    """
    if '${' not in value:
        return value
    encoded = base64.b64encode(value.encode('utf-8')).decode('ascii')
    return "${__groovy(new String(java.util.Base64.getDecoder().decode('" + encoded + "')\\, java.nio.charset.StandardCharsets.UTF_8))}"


def prop(parent, name, value, kind='stringProp'):
    element = ET.SubElement(parent, kind, name=name)
    element.text = str(value).lower() if isinstance(value, bool) else str(value)
    return element


def component(tree, tag, gui, name):
    element = ET.SubElement(tree, tag, guiclass=gui, testclass=tag, testname=name, enabled='true')
    children = ET.SubElement(tree, 'hashTree')
    return element, children


def write_plan(path, requests, threads=1, loops=1):
    root = ET.Element('jmeterTestPlan', version='1.2', properties='5.0', jmeter='5.6.3')
    top = ET.SubElement(root, 'hashTree')
    plan, plan_tree = component(top, 'TestPlan', 'TestPlanGui', 'REST API test plan')
    prop(plan, 'TestPlan.functional_mode', False, 'boolProp')
    prop(plan, 'TestPlan.serialize_threadgroups', True, 'boolProp')
    arguments = ET.SubElement(plan, 'elementProp', name='TestPlan.user_defined_variables', elementType='Arguments',
                              guiclass='ArgumentsPanel', testclass='Arguments', testname='User Defined Variables', enabled='true')
    ET.SubElement(arguments, 'collectionProp', name='Arguments.arguments')
    group, group_tree = component(plan_tree, 'ThreadGroup', 'ThreadGroupGui', 'API requests')
    prop(group, 'ThreadGroup.on_sample_error', 'continue')
    controller = ET.SubElement(group, 'elementProp', name='ThreadGroup.main_controller', elementType='LoopController',
                               guiclass='LoopControlPanel', testclass='LoopController', testname='Loop Controller', enabled='true')
    prop(controller, 'LoopController.continue_forever', False, 'boolProp')
    prop(controller, 'LoopController.loops', loops)
    prop(group, 'ThreadGroup.num_threads', threads)
    prop(group, 'ThreadGroup.ramp_time', '0')
    prop(group, 'ThreadGroup.scheduler', False, 'boolProp')
    prop(group, 'ThreadGroup.same_user_on_next_iteration', True, 'boolProp')

    for request in requests:
        sampler, sampler_tree = component(group_tree, 'HTTPSamplerProxy', 'HttpTestSampleGui', request.name)
        args = ET.SubElement(sampler, 'elementProp', name='HTTPsampler.Arguments', elementType='Arguments',
                             guiclass='HTTPArgumentsPanel', testclass='Arguments', testname='Request body', enabled='true')
        values = ET.SubElement(args, 'collectionProp', name='Arguments.arguments')
        if request.body_type != 'none':
            arg = ET.SubElement(values, 'elementProp', name='', elementType='HTTPArgument')
            prop(arg, 'HTTPArgument.always_encode', False, 'boolProp')
            prop(arg, 'Argument.name', '')
            prop(arg, 'Argument.value', literal(request.body))
            prop(arg, 'Argument.metadata', '=')
        # An absolute URL prevents JMeter from moving POST query parameters into the body.
        for key, value in [('HTTPSampler.domain', ''), ('HTTPSampler.port', ''), ('HTTPSampler.protocol', ''),
                           ('HTTPSampler.path', literal(request.url)), ('HTTPSampler.method', request.method),
                           ('HTTPSampler.contentEncoding', 'UTF-8'), ('HTTPSampler.implementation', 'HttpClient4'),
                           ('HTTPSampler.connect_timeout', request.connect_timeout_ms),
                           ('HTTPSampler.response_timeout', request.response_timeout_ms)]:
            prop(sampler, key, value)
        for key, value in [('HTTPSampler.follow_redirects', False), ('HTTPSampler.auto_redirects', False),
                           ('HTTPSampler.use_keepalive', True), ('HTTPSampler.DO_MULTIPART_POST', False),
                           ('HTTPSampler.image_parser', False), ('HTTPSampler.postBodyRaw', request.body_type != 'none')]:
            prop(sampler, key, value, 'boolProp')
        if request.headers:
            manager, _ = component(sampler_tree, 'HeaderManager', 'HeaderPanel', 'Request headers')
            collection = ET.SubElement(manager, 'collectionProp', name='HeaderManager.headers')
            for name, value in request.headers:
                header = ET.SubElement(collection, 'elementProp', name='', elementType='Header')
                prop(header, 'Header.name', literal(name))
                prop(header, 'Header.value', literal(value))
        assertion, _ = component(sampler_tree, 'ResponseAssertion', 'AssertionGui', 'Expected HTTP status')
        tests = ET.SubElement(assertion, 'collectionProp', name='Asserion.test_strings')
        prop(tests, '0', request.status_pattern)
        prop(assertion, 'Assertion.test_field', 'Assertion.response_code')
        prop(assertion, 'Assertion.assume_success', True, 'boolProp')
        prop(assertion, 'Assertion.test_type', 1, 'intProp')
        prop(assertion, 'Assertion.custom_message', 'HTTP status did not match expected_status')

    collector, _ = component(plan_tree, 'ResultCollector', 'SimpleDataWriter', 'Write results.jtl (CSV)')
    prop(collector, 'ResultCollector.error_logging', False, 'boolProp')
    config = ET.SubElement(ET.SubElement(collector, 'objProp'), 'name')
    config.text = 'saveConfig'
    value = ET.SubElement(collector.find('objProp'), 'value', attrib={'class': 'SampleSaveConfiguration'})
    enabled = ['time', 'latency', 'timestamp', 'success', 'label', 'code', 'message', 'threadName',
               'dataType', 'fieldNames', 'bytes', 'sentBytes', 'threadCounts', 'idleTime', 'connectTime',
               'saveAssertionResultsFailureMessage']
    disabled = ['encoding', 'assertions', 'subresults', 'responseData', 'samplerData', 'xml',
                'responseHeaders', 'requestHeaders', 'responseDataOnError', 'url']
    for name in enabled + disabled:
        ET.SubElement(value, name).text = 'true' if name in enabled else 'false'
    ET.SubElement(value, 'assertionsResultsToSave').text = '0'
    prop(collector, 'filename', '${__P(rest.results.file,results.jtl)}')
    listener, _ = component(plan_tree, 'JSR223Listener', 'TestBeanGUI', 'Log each API result')
    prop(listener, 'scriptLanguage', 'groovy')
    prop(listener, 'parameters', '')
    prop(listener, 'filename', '')
    prop(listener, 'cacheKey', 'true')
    prop(listener, 'script', '''def r = sampleResult
log.info('API_RESULT label={} code={} elapsedMs={} success={}',
    r.getSampleLabel(), r.getResponseCode(), r.getTime(), r.isSuccessful())
''')
    ET.indent(root, space='  ')
    path.write_bytes(xml_bytes(root) + b'\n')


RUNTIME_PROPERTIES = '''# Required execution settings for the generated JMeter 5.6.3 plan.
httpclient4.retrycount=0
httpclient4.request_sent_retry_enabled=false
httpclient4.idletimeout=0
http.java.sampler.retries=0
jmeter.save.saveservice.output_format=csv
jmeter.save.saveservice.default_delimiter=,
jmeter.save.saveservice.print_field_names=true
jmeter.save.saveservice.timestamp_format=ms
jmeter.save.saveservice.autoflush=true
sampleresult.default.encoding=UTF-8
summariser.name=summary
summariser.out=true
summariser.log=true
jmeterengine.nongui.port=0
'''
