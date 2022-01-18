/* eslint-disable */
import "../../assets/todoist-16.png"
import "../../assets/todoist-32.png"
import "../../assets/todoist-80.png"
import "../../assets/gantt-16.png"
import "../../assets/gantt-32.png"
import "../../assets/gantt-80.png"

import alert from "../components/alert.js"
import status from "../components/status.js"
import connect from "../components/connect.js"
import settings from "../components/settings.js"

let view = {

	state: {
		addLinks: false,
		avatar: "../../assets/avatar.png",
		pageTitle: "Onenote",
		pageUrl: location.href,
		tasks: [],
	},

	listeners: ["click", "change"],

	get(id) {
		return document.getElementById(id)
	},

	getValue(id) {
		let el = this.get(id)
		return el.value || el.innerText || ""
	},

	update(data) {
		for (let id in data) {
			let el = this.get(id)
			switch (el.nodeName) {
				case 'INPUT':
					el.value = data[id]
					break
				case 'IMG':
					el.src = data[id]
					break
				default:
					el.innerText = data[id]
			}
		}
	},
 
	show(component, data) {
		this.get("app-body").innerHTML = component.template

		this.listeners.forEach( action => {
			let elements = this.get('app-body').querySelectorAll(`[${action}]`)
			elements.forEach( el => {
				el.addEventListener(action, this[el.getAttribute(action)])
			})
		})

		let toggles = document.querySelectorAll(".ms-Toggle")
		toggles.forEach( toggle => new fabric['Toggle'](toggle) )

		if (data != undefined) this.update(data)
		return this
	},

	alert(title, details = "unknown") {
		this.show(alert, { title, details })
	},

	connect() {
		let token = view.getValue("token")
		if (token.length > 0) todoist.sync(token)
	},

	disconnect() {
		localStorage.removeItem('todoist_token')
		todoist.token = 'none'
		view.close()
	},

	toggle() {
		view.state.addLinks = !view.state.addLinks
	},

	refresh() {
		let projects = view.get("projects"),
		title = projects.selectedOptions[0].text,
		tasks = view.state.tasks.length + " new task(s)"

		if (projects.value == "new") {
			title = view.state.pageTitle
			view.update({ title, tasks })
		
		} else {
			todoist.getData(projects.value).then(data => {
				let items = [], todos = []
				data.items.forEach( item => items.push(tools.getText(item.content)) )
				todos = view.state.tasks.filter( todo => !items.includes(todo) )
				tasks = todos.length + " new task(s)"
				view.update({ title, tasks })
			})
		}
			
	},

	push() {
		let projectId = view.getValue("projects")

		if (projectId == "new") {
			todoist.push(view.state.tasks)

		} else {
			todoist.getData(projectId).then(data => {
				let items = [], tasks = []
				data.items.forEach( item => items.push(tools.getText(item.content)) )
				tasks = view.state.tasks.filter( todo => !items.includes(todo) )
				todoist.push(tasks, projectId)
			})
		}
	},

	retry() {
		getPageTasks()
	},

	close() {
		Office.context.ui.closeContainer()
	}

},

todoist = {

	token: localStorage["todoist_token"] || "none",
	url: "https://api.todoist.com/sync/v8/sync",

	sync(token) {
		if (token) this.token = token

		let headers = {
			'Authorization': 'Bearer ' + this.token,
			'Content-Type': 'application/json',
		},

		params = {
			sync_token: '*',
			resource_types: ["user", "projects"]
		},
		
		message = "Connecting..."
		view.show(status, { message })

		fetch(todoist.url, { 
			method: 'POST',
			headers : headers,
			body: JSON.stringify(params)
		})
	
		.then(res => {
			res.json().then(data => {
				localStorage.setItem("todoist_token", todoist.token)

				view.show(settings, {
					avatar: data.user.avatar_medium || view.state.avatar,
					user: data.user.full_name,
					mail: data.user.email.toLowerCase(),
					tasks: view.state.tasks.length + " task(s)",
					title: view.state.pageTitle,
					preview: view.state.tasks[0].substring(0, 34),
				})

				let list = view.get("projects")
				data.projects.forEach(project => {
					var selected = project.name == view.state.pageTitle ? true : false
					list.options.add(new Option(project.name, project.id, selected, selected))
				})

				let projects = view.get("dropdown")
				new fabric['Dropdown'](projects)

				projects.querySelectorAll('.ms-Dropdown-title')[0].innerHTML = list.selectedOptions[0].text
				view.refresh()

				if (view.state.addLinks) {
					let links = view.get("add-links")
					if (links != null) links.classList.add("is-selected")
				}
			})
		})

		.catch(error => {
			view.alert("Connection failed!", error);
		})
	},

	push(tasks = [], id = "new") {

		if (tasks.length > 0) {

			let project = { id, tasks },
			project_id = parseInt(project.id) || todoist.uuid(),
			commands = [],

			headers = {
				'Authorization': 'Bearer ' + this.token,
				'Content-Type': 'application/json'
			},

			message = "Proccesing..."
			view.show(status, { message })

			if (project.id == "new") {
				commands.push({
					type: "project_add",
					temp_id: project_id,
					uuid: todoist.uuid(),
					args: { name: view.state.pageTitle }
				})
			} 

			tasks.forEach( todo => {
				let link = view.state.pageUrl,
				content = view.state.addLinks ? `[${todo}](${link})` : todo
				commands.push({
					type: "item_add",
					temp_id: todoist.uuid(),
					uuid: todoist.uuid(),
					args: { content, project_id }
				})
			})
			
			fetch(todoist.url, { 
				method: 'POST',
				headers : headers,
				body: JSON.stringify({ commands })
			})
			
			.then(res => {
				res.json().then(data => {
					let projectId = data.temp_id_mapping[project_id] || project_id
					window.open("https://todoist.com/showProject?id=" + projectId, "_blank")
					view.close()
				})
			})

			.catch(error => {
				view.alert("Export failed!", error);
			})

		} else {
			view.alert("No new tasks found.", 
					   "No new tasks found on this page!")
		}

	},

	uuid() {
		return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
			let d = new Date().getTime(),
			r = (d + Math.random()*16)%16 | 0
			d = Math.floor(d/16)
			return (c=='x' ? r : (r&0x7|0x8)).toString(16)
		})
	},

	async getData(project_id) {
		return await fetch("https://api.todoist.com/sync/v8/projects/get_data", { 
			method: 'POST',
			headers : {
				'Authorization': 'Bearer ' + this.token,
				'Content-Type': 'application/json'
			},
			body: JSON.stringify({ project_id })
		})
		.then( res => res.json())
		.catch(error => {
			view.alert("Todoist sync failed!", error);
		})
	},

},

tools = {
	getText(content) {
		return content.split(']')[0].replace(/\s/gi, " ").replace(/\[/gi, "")
	},
}

Office.onReady((info) => {
	if (info.host === Office.HostType.OneNote) getPageTasks()
})

export async function getPageTasks() {

	view.state.tasks = []
	view.show(status)
	OneNote.run(context => {

		let parser = new DOMParser(),
		page = context.application.getActivePage(),
		outlines = [],
		tables = [],
		strings = []

		page.load("title, webUrl")
		page.contents.load("type, items");
		return context.sync()

		.then(() => {
			view.state.pageTitle = page.title
			view.state.pageUrl = page.webUrl
			page.contents.items.forEach(item => {
				if (item.type == 'Outline') {
					item.load("outline/paragraphs, outline/paragraphs/type")
					outlines.push(item)
				}
			})
			return context.sync()
		})

		.then(() => {
			outlines.forEach(item => {
				item.outline.paragraphs.items.forEach( p => {
					if (p.type == "RichText") {
						p.load("richtext")
						strings.push(p.richText.getHtml())
					} else if (p.type == "Table") {
						p.load("table/rows/items/cells/paragraphs/type") 
						tables.push(p.table)
					}
				})
			})
			return context.sync()
		})

		.then(() => {
			tables.forEach( table => {
				table.rows.items.forEach( row => {
					row.cells.items.forEach( cell => {
						cell.paragraphs.items.forEach( p => {
							if (p.type == "RichText") {
								p.load("richtext")
								strings.push(p.richText.getHtml())
							}
						})
					})
				}) 
			})
			return context.sync()
		})

		.then(() => {
			strings.forEach( html => {
				let doc = parser.parseFromString(html.value, 'text/html'),
				tag = doc.querySelector("[data-tag=to-do]")
				if (tag != null) view.state.tasks.push(tag.innerText.trim())
			})
			console.log('tasks found:', view.state.tasks.length)
			
			if (view.state.tasks.length == 0)
				view.alert("No tasks found! There is nothing to export.", 
						   "No to-do tags found on this page!")
			else if (todoist.token == "none") view.show(connect)
			else todoist.sync()
			
		})
		
	})

	.catch(error => {
		view.alert("Error!", error);
	})

}
