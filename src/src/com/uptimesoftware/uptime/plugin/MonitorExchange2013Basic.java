package com.uptimesoftware.uptime.plugin;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.net.Socket;
import java.net.UnknownHostException;
import java.util.LinkedList;

import org.codehaus.jackson.JsonFactory;
import org.codehaus.jackson.JsonNode;
import org.codehaus.jackson.JsonParser;
import org.codehaus.jackson.map.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import ro.fortsoft.pf4j.PluginWrapper;
import com.uptimesoftware.uptime.plugin.api.Extension;
import com.uptimesoftware.uptime.plugin.api.Plugin;
import com.uptimesoftware.uptime.plugin.api.PluginMonitor;
import com.uptimesoftware.uptime.plugin.monitor.MonitorState;
import com.uptimesoftware.uptime.plugin.monitor.Parameters;

/**
 * Exchange2013 Basic Monitor / Exchange 2013.
 * 
 * @author uptime software
 */
public class MonitorExchange2013Basic extends Plugin {

	/**
	 * Constructor - a plugin wrapper.
	 * 
	 * @param wrapper
	 */
	public MonitorExchange2013Basic(PluginWrapper wrapper) {
		super(wrapper);
	}

	/**
	 * A nested static class which has to extend PluginMonitor.
	 * 
	 * Functions that require implementation :
	 * 1) The monitor function will implement the main functionality and should set the monitor's state and result
	 * message prior to completion.
	 * 2) The setParameters function will accept a Parameters object containing the values filled into the monitor's
	 * configuration page in Up.time.
	 */
	@Extension
	public static class UptimeMonitorExchange2013Basic extends PluginMonitor {
		// Logger object.
		private static final Logger LOGGER = LoggerFactory.getLogger(UptimeMonitorExchange2013Basic.class);

		// Constants.
		private static final String VERSION = "ver";
		private static final String OS_TYPE_WINDOWS = "Windows";
		private static final String SCRIPT_COMMAND_NAME = "exchange2013monitor";
		private static final String JSON_ATTRIBUTE_RESULT = "result";
		private static final int ERROR_CODE = -1;
		private static final String AGENT_ERR = "ERR";

		// 12 outputs for now, may increase in the future.
		private final LinkedList<String> outputList = new LinkedList<String>();

		// See definition in .xml file for plugin. Each plugin has different number of input/output parameters.
		// [Input]
		String hostname = "";
		int port = 0;
		String password = "";

		/**
		 * The setParameters function will accept a Parameters object containing the values filled into the monitor's
		 * configuration page in Up.time.
		 * 
		 * @param params
		 *            Parameters object which contains inputs.
		 */
		@Override
		public void setParameters(Parameters params) {
			LOGGER.debug("Step 1 : Setting parameters.");
			// [Input]
			hostname = params.getString("hostname");
			port = params.getInt("port");
			password = params.getString("password");
			// [Outputs] in ArrayList
			// Transport Smtp Receive
			outputList.add("SMTPAverageBytesPerInboundMessage");
			outputList.add("SMTPBytesReceivedPerSecond");
			outputList.add("SMTPInboundConnections");
			outputList.add("SMTPInboundMessagesReceivedPerSecond");
			// Transport Smtp Send
			outputList.add("SMTPBytesSentPerSecond");
			outputList.add("SMTPOutboundConnections");
			outputList.add("SMTPMessagesSentPerSecond");
			// OWA
			outputList.add("CurrentWebmailUsers");
			outputList.add("WebmailUserLogonsPerSecond");
			// RPC
			outputList.add("RPCAveragedLatency");
			outputList.add("RPCClientsBytesRead");
			outputList.add("RPCClientsBytesWritten");
			outputList.add("RPCDispatchTaskActiveThreads");
			outputList.add("RPCDispatchTaskQueueLength");
			outputList.add("RPCOperationsPersec");
			outputList.add("RPCRequests");
			// Assistants - Per Assistant
			outputList.add("AverageEventProcessingTimeInSecondsPerAssistant");
			outputList.add("AverageEventQueueTimeInSecondsPerAssistant");
			outputList.add("EventsInQueuePerAssistant");
			outputList.add("EventsProcessedPerAssistant");
			outputList.add("EventsProcessedPerSecondPerAssistant");
			// Assistants - Per Database
			outputList.add("AverageEventProcessingTimeInSecondsPerDatabase");
			outputList.add("AverageMailboxProcessingTimeInSecondsPerDatabase");
			outputList.add("EventsInQueuePerDatabase");
			outputList.add("MailboxesProcessedPerDatabase");
			outputList.add("MailboxesProcessedPerSecondPerDatabase");
		}

		/**
		 * The monitor function will implement the main functionality and should set the monitor's state and result
		 * message prior to completion.
		 */
		@Override
		public void monitor() {
			LOGGER.debug("Step 2 : Send a command to Up.time Agent and get a name of current OS.");
			String osType = sendCmdToAgent(hostname, port, VERSION);

			LOGGER.debug("Error handling : Check if Up.time Agent sent back ERR message");
			if (osType.equals(AGENT_ERR)) {
				setStateAndMessage(MonitorState.CRIT, "Agent sent back ERR. Check the ver command.");
				// plugin stops.
				return;
			}

			LOGGER.debug("Error handling : Check if Up.time Agent sent back OS info");
			if (osType == null | osType.equals("")) {
				setStateAndMessage(MonitorState.CRIT, "Could not get OS info from Up.time Agent");
				return;
			}

			LOGGER.debug("Step 3 : Verify that the plugin is running on Windows. If it's running on Linux, ERROR");
			if (!osType.contains(OS_TYPE_WINDOWS)) {
				setStateAndMessage(MonitorState.CRIT, "Exhange 2013 plugin cannot run on Linux. (Windows Only)");
				return;
			}

			LOGGER.debug("Step 4 : Plugin is running on Windows. Send rexec command to run a custom script");
			String JSONResult = runWindowsCustomScript(hostname, port, password, SCRIPT_COMMAND_NAME);

			LOGGER.debug("Error handling : Check if Up.time Agent sent back ERR message");
			if (JSONResult.equals(AGENT_ERR)) {
				setStateAndMessage(MonitorState.CRIT, "Agent sent back ERR. Check the rexec command.");
				return;
			}

			LOGGER.debug("Error handling : Check if Up.time Agent sent back OS info");
			if (JSONResult == null | JSONResult.equals("")) {
				setStateAndMessage(MonitorState.CRIT, "Could not get JSON result from Up.time Agent");
				return;
			}

			LOGGER.debug("Step 5 : Convert result String in JSON format to JsonNode object.");
			JsonNode nodeObject = convertStringToJsonNode(JSONResult);

			LOGGER.debug("Error handling : Check if JsonNode object is null or not.");
			if (nodeObject == null) {
				setStateAndMessage(MonitorState.CRIT, "Could not convert result String to JsonNode object");
				return;
			}

			LOGGER.debug("Step 6 : Parse output values from JsonNode object.");
			int outputValue = 0;
			String outputParam = "";
			while (outputList.size() != 0) {
				outputParam = outputList.pop();
				outputValue = getIntValueFromJsonNode(outputParam, nodeObject);
				if (outputValue == ERROR_CODE) {
					setStateAndMessage(MonitorState.CRIT, "Unable to get int value of " + outputParam);
					return;
				} else {
					addVariable(outputParam, outputValue);
				}
			}

			LOGGER.debug("Step 7 : Everything ran okay. Set monitor state to OK");
			setStateAndMessage(MonitorState.OK, "Monitor successfully ran.");
		}

		/**
		 * Private helper function to get int value of specified field.
		 * 
		 * @param fieldName
		 *            Name of field that has int value.
		 * @param nodeObject
		 *            JsonNode object containing all JSON nodes.
		 * @return int value of specified field.
		 */
		private int getIntValueFromJsonNode(String fieldName, JsonNode nodeObject) {
			// Get nested JSON.
			JsonNode nestedJsonNode = nodeObject.get(JSON_ATTRIBUTE_RESULT);
			if (nestedJsonNode == null) {
				LOGGER.error("Could not find {} attribute.", JSON_ATTRIBUTE_RESULT);
				return ERROR_CODE;
			}
			nestedJsonNode = nestedJsonNode.get(fieldName);
			if (nestedJsonNode == null) {
				LOGGER.error("Could not find {} attribute within {} attribute", fieldName, JSON_ATTRIBUTE_RESULT);
				return ERROR_CODE;
			}
			return Integer.parseInt(nestedJsonNode.getTextValue());
		}

		/**
		 * Private helper function to convert result String in JSON format to JsonNode object.
		 * (3rd party library - Jackson)
		 * 
		 * @param JSONFormatResult
		 *            Result String in JSON format.
		 * @return JsonNode object containing all JSON nodes.
		 */
		private JsonNode convertStringToJsonNode(String JSONFormatResult) {
			ObjectMapper mapper = new ObjectMapper();
			JsonFactory factory = mapper.getJsonFactory();
			JsonParser JSONParser = null;
			JsonNode nodeObject = null;
			try {
				JSONParser = factory.createJsonParser(JSONFormatResult);
				nodeObject = mapper.readTree(JSONParser);
			} catch (IOException e) {
				LOGGER.error("Error while reading JSON tree", e);
			}
			if (nodeObject == null) {
				LOGGER.error("Converting did not complete successfully.");
			}
			return nodeObject;
		}

		/**
		 * Private helper function to run a custom script on Windows.
		 * 
		 * @param host
		 *            Name of host
		 * @param port
		 *            Port number
		 * @param cmd
		 *            Command to run a custom script on Windows
		 * @return Response from a socket in String.
		 */
		private String runWindowsCustomScript(String host, int port, String password, String cmd) {
			return sendCmdToAgent(host, port, "rexec " + password + " " + cmd);
		}

		/**
		 * Private helper function to open a socket and write to & read from the open socket.
		 * 
		 * @param host
		 *            Name of host
		 * @param port
		 *            Port number
		 * @param cmd
		 *            Command to write to a open socket
		 * @return Response from a socket in String.
		 */
		private String sendCmdToAgent(String host, int port, String cmd) {
			StringBuilder result = new StringBuilder();
			// Try-with-resource statement will close socket and resources after completing a task.
			try (Socket socket = new Socket(host, port);
					BufferedWriter out = new BufferedWriter(new OutputStreamWriter(socket.getOutputStream()));
					BufferedReader in = new BufferedReader(new InputStreamReader(socket.getInputStream()));) {
				// Check if cmd is empty or not.
				if (cmd.equals("") || cmd == null) {
					LOGGER.error("{} is empty/null", cmd);
				} else {
					// Write the cmd on the connected socket.
					out.write(cmd);
					// Flush OutputStream after writing because there is no guarantee that the serialized representation
					// will get sent to the other end.
					out.flush();
				}
				// Read line(s) from the connected socket.
				String line = "";
				while ((line = in.readLine()) != null) {
					result.append(line);
				}
			} catch (UnknownHostException e) {
				LOGGER.error("Unable to open a socket.", e);
			} catch (IOException e) {
				LOGGER.error("Unable to get I/O of socket.", e);
			}
			return result.toString();
		}
	}
}