package com.priceline.finance.attachmentToDisk;

import java.net.URI;
import java.util.ArrayList;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.FolderTraversal;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.FolderSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

/*
 * @author Manas Gupta
 * @since 21 Oct, 2015
 */
public class EWSAttachmentToDisk {

	public static void main(String[] args) throws Exception {
		ExchangeService service = new ExchangeService(
				ExchangeVersion.Exchange2007_SP1);
		ExchangeCredentials credentials = new WebCredentials(args[0], args[1],
				"CORP");
		service.setCredentials(credentials);
		service.setUrl(new URI(
				"https://nw-mail.corp.priceline.com/EWS/Exchange.asmx"));

		FolderView folderView = new FolderView(100);
		folderView.setPropertySet(new PropertySet(BasePropertySet.IdOnly));
		folderView.getPropertySet().add(FolderSchema.DisplayName);
		folderView.getPropertySet().add(FolderSchema.ChildFolderCount);
		folderView.setTraversal(FolderTraversal.Deep);
		FindFoldersResults findFolderResults = service.findFolders(
				WellKnownFolderName.Root, folderView);
		ArrayList<Folder> listOfFolders = findFolderResults.getFolders();
		for (Folder f : listOfFolders) {
			if (f.getDisplayName().equalsIgnoreCase("Finance Support")) {
				System.out.println("Name of Folder : " + f.getDisplayName());
				ItemView view = new ItemView(100);
				SearchFilter unreadFilter = new SearchFilter.SearchFilterCollection(
						LogicalOperator.And, new SearchFilter.IsEqualTo(
								EmailMessageSchema.IsRead, false));
				FindItemsResults findResults = service.findItems(f.getId(),
						unreadFilter, view);
				ArrayList<Item> listOfItems = findResults.getItems();
				System.out.println("Total Unread E-mails : "
						+ listOfItems.size());
				for (Item item : listOfItems) {
					item.load(new PropertySet(
							BasePropertySet.FirstClassProperties,
							ItemSchema.MimeContent));
					System.out.println("Subject : " + item.getSubject());
					// Need to set item as READ.
					item.update(ConflictResolutionMode.AutoResolve);
				}
				f.update();
			}
		}
	}
}