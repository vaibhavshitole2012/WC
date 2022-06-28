package ext.kc.product.createbom;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.rmi.RemoteException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import wt.epm.EPMDocument;
import wt.epm.build.EPMBuildRule;
import wt.epm.structure.EPMDescribeLink;
import wt.fc.PersistenceHelper;
import wt.fc.QueryResult;
import wt.fc.WTObject;
import wt.folder.Folder;
import wt.folder.FolderHelper;
import wt.httpgw.GatewayAuthenticator;
import wt.inf.container.WTContainerHelper;
import wt.inf.container.WTContainerRef;
import wt.lifecycle.LifeCycleHelper;
import wt.lifecycle.LifeCycleTemplate;
import wt.method.RemoteAccess;
import wt.method.RemoteMethodServer;
import wt.part.PartDocHelper;
import wt.part.WTPart;
import wt.part.WTPartHelper;
import wt.part.WTPartMaster;
import wt.part.WTPartUsageLink;
import wt.pds.StatementSpec;
import wt.query.QuerySpec;
import wt.query.SearchCondition;
import wt.query.WhereExpression;
import wt.util.WTException;
import wt.util.WTPropertyVetoException;
import wt.vc.VersionControlHelper;
import wt.vc.config.ConfigSpec;
import wt.vc.config.LatestConfigSpec;
import wt.vc.wip.WorkInProgressHelper;

/** 
 * windchill ext.kc.product.createbom.CreateBOMStructure <admin_user_id> <input
 * file path>
 */
public class CreateBOMStructure2 implements RemoteAccess {

	private static final String CLASSNAME = CreateBOMStructure2.class.getName();

	static List<String> safetyStockList;
	static List<String> gradeChangePartsList;
	static List<String> itemList;
	static List<String> qtyList;
	static List<String> orderdescritpionList;
	static List<String> drawingNumberList;
	static List<Integer> kcNumberList;
	//static List<Integer> kcProductList;

	public static void main(String[] args) {

		if (args.length < 2 || args.length > 2) {
			System.out.println("Wrong Number of Argument Parameters.");
		}
		String username = null;
		String inputFile = null;

		try {
			username = args[0];
			inputFile = args[1];
			System.out.println("UserName : " + username + "\nInputFile : " + inputFile);
		} catch (Exception ex) {
			// printUsage();
			throw new ArrayIndexOutOfBoundsException("Invalid input param");
		}

		Class[] aClass = { String.class };
		Object[] aObj = { inputFile };

		RemoteMethodServer rms = RemoteMethodServer.getDefault();
		GatewayAuthenticator auth = new GatewayAuthenticator();
		auth.setRemoteUser(username);
		rms.setAuthenticator(auth);

		try {
			System.out.println("Utility Started >>>> : " + java.time.Clock.systemUTC().instant());
			rms.invoke("execute", CLASSNAME, null, aClass, aObj);
			System.out.println("Utility Executed successfully >>>> : " + java.time.Clock.systemUTC().instant());
		} catch (InvocationTargetException | RemoteException e) {

			// LOG.error(e.getLocalizedMessage(), e);
			System.out.println("Unable to invoke " + e);
		}
	}



	public static void execute(String filePath) throws WTException, IOException, WTPropertyVetoException {
		readFromFile(filePath);
		retrievePartFromList();
	}


	private static void readFromFile(String path) throws IOException {

		int i = 0;
		safetyStockList = new ArrayList<>();
		gradeChangePartsList = new ArrayList<>();
		itemList = new ArrayList<>();
		qtyList = new ArrayList<>();
		orderdescritpionList = new ArrayList<>();
		drawingNumberList = new ArrayList<>();
		kcNumberList = new ArrayList<>();
		//kcProductList = new ArrayList<>();

		try {
			File file = new File("D:\\SampleBOM\\BOM.xlsx"); // creating // // instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
			// creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator(); // iterating over excel file
			Row headerRow = itr.next(); // Skip Header
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// cell.setCellType(Cell.CELL_TYPE_STRING);

					i = cell.getColumnIndex();
					if (i == 0) {
						String safetyStock = cell.getStringCellValue();
						safetyStockList.add(safetyStock.trim());
						System.out.println("safetyStock......." + safetyStock);
						System.out.println("safetyStockList......." + safetyStockList);
					}
					if (i == 1) {
						String gradeChangeParts = cell.getStringCellValue();
						gradeChangePartsList.add(gradeChangeParts.trim());

						System.out.println("gradeChangeParts......." + gradeChangeParts);
						System.out.println("gradeChangePartsList......." + gradeChangePartsList);
					}
					if (i == 2) {
						
						// String item = NumberToTextConverter.toText(cell.getNumericCellValue());
						String item = String.valueOf(cell.getNumericCellValue());
						itemList.add(item);
						System.out.println("item......." + item);
						System.out.println("itemList......." + itemList);

					}
					if (i == 3) {
						String qty = String.valueOf(cell.getNumericCellValue());
						qtyList.add(qty);
						System.out.println("qty......." + qty);
						System.out.println("qtyList......." + qtyList);
					}
					if (i == 4) {

						String drawingNumber = cell.getStringCellValue();
						drawingNumberList.add(drawingNumber);

						System.out.println("drawingNumber......." + drawingNumber);
						System.out.println("drawingNumberList......." + drawingNumberList);
					}
					if (i == 5) {

						Integer kcNumber = (int) cell.getNumericCellValue();
						kcNumberList.add(kcNumber);

						System.out.println("kcNumber......." + kcNumber);
						System.out.println("kcNumberList......." + kcNumberList);
					}
//					if (i == 6) {
//
//						Integer kcproduct = (int) cell.getNumericCellValue();
//						kcProductList.add(kcproduct);
//
//						System.out.println("kcNumber......." + kcproduct);
//						System.out.println("kcProductList......." + kcProductList);
//					}

					if (i == 6) {
						String orderdescitpion = cell.getStringCellValue();
						orderdescritpionList.add(orderdescitpion);

						System.out.println("orderdescitpion......." + orderdescitpion);
						System.out.println("orderdescritpionList......." + orderdescritpionList);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	private static void retrievePartFromList() throws RemoteException, WTPropertyVetoException {

		String partNum = null;
		String quantity = null;
		WTPart part = null;
		WTPart parentPart = null;
		try {
			for (int index = 0; index < itemList.size(); index++) {
				if (index != 0) {
					parentPart = getPart(drawingNumberList.get(0));
					partNum = drawingNumberList.get(index);
					quantity = qtyList.get(index);

					part = getPart(partNum);

					if (part != null) {
						bomProcessed(parentPart, part, quantity, partNum);
					} else {
						createPart(partNum);
						part = getPart(partNum);
						if (part != null) {
							bomProcessed(parentPart, part, quantity, partNum);
						}
					}
				} else {
					partNum = drawingNumberList.get(0);
					part = getPart(partNum);
					if (part != null) {
						continue;
					} else {
						createPart(partNum);
					}
				}
			}
		} catch (WTException e) {
			e.printStackTrace();
		}
	}

	private static void bomProcessed(WTPart parentPart, WTPart part, String quantity, String epmNum)
			throws WTException, WTPropertyVetoException {
		if (!checkWhereUsed(part)) {
			createBOM(parentPart, part, quantity, epmNum);
		} else {
			System.out.println("Part is already associated");
		}
	}

	private static boolean checkWhereUsed(WTPart currentPart) throws WTException {
		System.out.println("Inside checkWhereUsed");
		QueryResult qr = WTPartHelper.service.getUsedByWTParts((WTPartMaster) currentPart.getMaster());
		System.out.println("Part is used: " + qr.size());
		if (qr.size() != 0) {
			return true;
		} else {
			return false;
		}
	}

	private static void createBOM(WTPart parentPart2, WTPart currentPart, String quantityValue, String epmNum) {
		System.out.println("Inside createBOM Method");
		WTPart parentPart = null;
		WTPart childPart = null;
		WTPartMaster master = null;

		try {
			parentPart = (WTPart) VersionControlHelper.service.getLatestIteration(parentPart2, false);
			childPart = (WTPart) VersionControlHelper.service.getLatestIteration(currentPart, false);

			parentPart = checkPartCheckout(parentPart);
			childPart = checkPartCheckout(childPart);
			//associateDrw(epmNum, childPart);
			//associatePrt(epmNum, childPart);
			master = (WTPartMaster) childPart.getMaster();

			WTPartUsageLink link = WTPartUsageLink.newWTPartUsageLink(parentPart, master);
			setQuantity(link, quantityValue);

			PersistenceHelper.manager.save(link);
			WorkInProgressHelper.service.checkin(childPart, "");
			WorkInProgressHelper.service.checkin(parentPart, "");
			System.out.println("Two parts are associated.");

		} catch (WTException | WTPropertyVetoException e) {
			e.printStackTrace();
		}
	}

	private static WTPart checkPartCheckout(WTPart part) {
		System.out.println("Inside checkPartCheckout");
		try {
			if (!WorkInProgressHelper.isCheckedOut(part)) {
				WorkInProgressHelper.service.checkout(part, WorkInProgressHelper.service.getCheckoutFolder(), "");
				part = (WTPart) WorkInProgressHelper.service.workingCopyOf(part);
				System.out.println(part.getNumber() + " Part is Checked out");
			} else {
				System.out.println(part.getNumber() + " Part is Already Checked out");
				part = (WTPart) WorkInProgressHelper.service.workingCopyOf(part);
			}
		} catch (WTException | WTPropertyVetoException e) {
			e.printStackTrace();
		}
		return part;
	}

	private static void associateDrw(String epmDocNum, WTPart part) throws WTException, WTPropertyVetoException {
		System.out.println("Inside associateDrw");
		WTPart childPart = null;
		List<String> numList = new ArrayList<String>();

		QueryResult associatedParts = PartDocHelper.service.getAssociatedDocuments(part);
		if (associatedParts.size() != 0) {
			while (associatedParts.hasMoreElements()) {
				WTObject obj = (WTObject) associatedParts.nextElement();
				if (obj instanceof EPMDocument) {
					EPMDocument epm = (EPMDocument) obj;
					numList.add(epm.getNumber());
				}
			}
		}
		EPMDocument epmDoc = getEPM(epmDocNum);

		if (epmDoc != null && numList.size() == 0 && !numList.contains(epmDoc.getNumber() + ".SLDDRW")) {
			childPart = (WTPart) VersionControlHelper.service.getLatestIteration(part, false);
			EPMDescribeLink describeLink = EPMDescribeLink.newEPMDescribeLink(childPart, epmDoc);
			describeLink = (EPMDescribeLink) PersistenceHelper.manager.save(describeLink);
			System.out.println(part.getNumber() + " Part is associated with EPMDocument " + epmDoc.getNumber());
		} else {
			System.out.println(epmDoc.getNumber() + "Drw is already associated with current Part " + part.getNumber());
		}
	}

	private static EPMDocument getEPM(String drwNum) throws WTException {
		System.out.println("Getting EPM: " + drwNum);
		drwNum = drwNum + ".SLDDRW";
		EPMDocument epm = null;
		QuerySpec querySpec = new QuerySpec(EPMDocument.class);
		WhereExpression searchCondition = new SearchCondition(EPMDocument.class, EPMDocument.NUMBER,
				SearchCondition.EQUAL, drwNum, false);
		querySpec.appendWhere(searchCondition, new int[] { 0 });
		ConfigSpec configSpec = new LatestConfigSpec();
		querySpec = configSpec.appendSearchCriteria(querySpec);

		QueryResult result = PersistenceHelper.manager.find((StatementSpec) querySpec);

		if (result.hasMoreElements()) {
			epm = (EPMDocument) result.nextElement();
		}
		return epm;
	}

	private static void associatePrt(String epmPrtNum, WTPart part) throws WTException {
		System.out.println("Inside associatePrt");
		WTPart childPart = null;
		List<String> numList = new ArrayList<String>();

		QueryResult associatedParts = PartDocHelper.service.getAssociatedDocuments(part);
		if (associatedParts.size() != 0) {
			while (associatedParts.hasMoreElements()) {
				WTObject obj = (WTObject) associatedParts.nextElement();
				if (obj instanceof EPMDocument) {
					EPMDocument epm = (EPMDocument) obj;
					numList.add(epm.getNumber());
				}
			}
		}
		EPMDocument epmPrtDoc = getEPMPrt(epmPrtNum);

		if (epmPrtDoc != null && numList.size() == 0 && !numList.contains(epmPrtDoc.getNumber() + ".SLDDRW")) {
			childPart = (WTPart) VersionControlHelper.service.getLatestIteration(part, false);
			EPMBuildRule buildRule = EPMBuildRule.newEPMBuildRule(epmPrtDoc, childPart,
					EPMBuildRule.BUILD_ATTRIBUTES | EPMBuildRule.BUILD_STRUCTURE | EPMBuildRule.CAD_REPRESENTATION);
			buildRule = (EPMBuildRule) PersistenceHelper.manager.save(buildRule);
			System.out.println(part.getNumber() + " Part is associated with EPMDocument " + epmPrtDoc.getNumber());
		} else {
			System.out
					.println(epmPrtDoc.getNumber() + "Prt is already associated with current Part " + part.getNumber());
		}
	}

	private static EPMDocument getEPMPrt(String epmPrtNum) throws WTException {
		System.out.println("Getting EPM: " + epmPrtNum);
		epmPrtNum = epmPrtNum + ".SLDPRT";
		EPMDocument epm = null;
		QuerySpec querySpec = new QuerySpec(EPMDocument.class);
		WhereExpression searchCondition = new SearchCondition(EPMDocument.class, EPMDocument.NUMBER,
				SearchCondition.EQUAL, epmPrtNum, false);
		querySpec.appendWhere(searchCondition, new int[] { 0 });
		ConfigSpec configSpec = new LatestConfigSpec();
		querySpec = configSpec.appendSearchCriteria(querySpec);

		QueryResult result = PersistenceHelper.manager.find((StatementSpec) querySpec);

		if (result.hasMoreElements()) {
			epm = (EPMDocument) result.nextElement();
		}
		return epm;
	}

	private static void setQuantity(WTPartUsageLink link, String quantityValue) {
		System.out.println("Inside setQuantity");
		wt.part.Quantity quantity = new wt.part.Quantity();
		if (!quantityValue.equalsIgnoreCase("null")) {
			quantity.setAmount(Double.parseDouble(quantityValue));
		} else {
			quantity.setAmount(Double.parseDouble("0"));
		}
		link.setQuantity(quantity);
	}

	@SuppressWarnings("deprecation")
	private static WTPart createPart(String partNum) {
		System.out.println("Creating a part: " + partNum);
		String lifeCycleName = "Basic";
		String container_path = "/wt.inf.container.OrgContainer=KIMBERLY-CLARK CORPORATION/wt.pdmlink.PDMLinkProduct=153P - Diapers Paris"; // The container
																										// where the
																										// document will
																										// be
																										// created/located
		String folder_path = "/Default";
		WTPart part = null;
		try {
			WTContainerRef containerRef = WTContainerHelper.service.getByPath(container_path);
			part = WTPart.newWTPart(partNum, "MyPartName");
			part.setContainer(containerRef.getReferencedContainer());
			Folder folder = FolderHelper.service.getFolder(folder_path, containerRef);
			FolderHelper.assignLocation(part, folder);
			LifeCycleTemplate lct = LifeCycleHelper.service.getLifeCycleTemplate(lifeCycleName,
					part.getContainerReference());
			part = (WTPart) LifeCycleHelper.setLifeCycle(part, lct);
			part = (WTPart) wt.fc.PersistenceHelper.manager.save(part);
			System.out.println(part.getNumber() + " Part is Created");
		} catch (WTPropertyVetoException | WTException e) {
			e.printStackTrace();
		}
		return part;

	}

	private static WTPart getPart(String partNum) throws WTException {
		System.out.println("Getting Part: " + partNum);
		WTPart part = null;
		QuerySpec querySpec = new QuerySpec(WTPart.class);
		WhereExpression searchCondition = new SearchCondition(WTPart.class, WTPart.NUMBER, SearchCondition.EQUAL,
				partNum, false);
		querySpec.appendWhere(searchCondition, new int[] { 0 });
		ConfigSpec configSpec = new LatestConfigSpec();
		querySpec = configSpec.appendSearchCriteria(querySpec);
		QueryResult result = PersistenceHelper.manager.find((StatementSpec) querySpec);

		if (result.hasMoreElements()) {
			part = (WTPart) result.nextElement();
		}
		return part;
	}

	private static void printUsage() {
		System.out.println("Usage:");
		System.out.println("\t - Windchill username");
		System.out.println("\t - Path to read file with input data");
	}

}